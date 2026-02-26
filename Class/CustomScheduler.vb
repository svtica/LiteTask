Imports System.Collections.Concurrent
Imports System.Data
Imports LiteTask.LiteTask.ScheduledTask


Namespace LiteTask
    Public Class CustomScheduler
        Implements IDisposable
        Public Event TaskCompleted As EventHandler(Of TaskCompletedEventArgs)

        Private _disposed As Boolean
        Private ReadOnly _credentialManager As CredentialManager
        Private ReadOnly _xmlManager As XMLManager
        Private ReadOnly _logger As Logger
        Private ReadOnly _toolManager As ToolManager
        Private ReadOnly _tasks As New ConcurrentDictionary(Of String, ScheduledTask)
        Private ReadOnly _errorNotifier As ErrorNotifier
        Private ReadOnly _taskRunner As TaskRunner
        Private disposedValue As Boolean
        Private ReadOnly _taskRunning As New ConcurrentDictionary(Of String, Boolean)
        Private ReadOnly _dependencyManager As TaskDependencyManager
        Private ReadOnly _taskLocks As New ConcurrentDictionary(Of String, Object)
        Private ReadOnly _mutexBasePath As String = "Global\LiteTask_Task_"
        Private ReadOnly _powerShellPathManager As PowerShellPathManager
        Private ReadOnly _taskStates As New ConcurrentDictionary(Of String, TaskState)
        Private ReadOnly _schedulerLock As New SemaphoreSlim(1, 1)
        Private Const STALE_TIMEOUT_MINUTES As Integer = 15
        Private ReadOnly _processingLock As New SemaphoreSlim(1, 1)
        Private _isProcessing As Boolean = False
        Private ReadOnly _staleTaskAlerts As New ConcurrentDictionary(Of String, DateTime)
        Private _lastStaleCleanupTime As DateTime = DateTime.MinValue
        Private Const STALE_CLEANUP_INTERVAL_SECONDS As Integer = 300

        ' Memory monitoring properties
        Private Const MEMORY_CHECK_INTERVAL_SECONDS As Integer = 300  ' Check every 5 minutes
        Private Const MEMORY_WARNING_THRESHOLD_MB As Long = 1024      ' Warn at 1 GB
        Private Const MEMORY_CRITICAL_THRESHOLD_MB As Long = 2048     ' Critical at 2 GB
        Private _lastMemoryCheckTime As DateTime = DateTime.MinValue
        Private _lastMemoryAlertTime As DateTime = DateTime.MinValue
        Private Const MEMORY_ALERT_COOLDOWN_MINUTES As Integer = 30   ' Avoid alert spam

        ' Enhanced mutex management properties
        Private Const MUTEX_TIMEOUT_SECONDS As Integer = 30  ' Increased from 1 second
        Private Const MAX_CLEANUP_RETRIES As Integer = 3
        Private ReadOnly _activeMutexes As New ConcurrentDictionary(Of String, DateTime)
        Private ReadOnly _mutexCleanupTimer As System.Threading.Timer


        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
        Private Structure TASK_INFO
            Public InstanceCount As UInteger
            Public LastRunTime As System.Runtime.InteropServices.ComTypes.FILETIME
            Public NextRunTime As System.Runtime.InteropServices.ComTypes.FILETIME
            Public Status As TASK_STATUS
            <MarshalAs(UnmanagedType.LPWStr)>
            Public ApplicationName As String
            <MarshalAs(UnmanagedType.LPWStr)>
            Public Parameters As String
            <MarshalAs(UnmanagedType.LPWStr)>
            Public WorkingDirectory As String
            <MarshalAs(UnmanagedType.LPWStr)>
            Public Creator As String
            <MarshalAs(UnmanagedType.LPWStr)>
            Public Comment As String
        End Structure

        Private Enum TASK_STATUS
            TASK_STATUS_READY = 2
            TASK_STATUS_RUNNING = 4
            TASK_STATUS_DISABLED = 8
        End Enum

        Public Sub New(credentialManager As CredentialManager, xmlManager As XMLManager,
              toolManager As ToolManager, logger As Logger, taskRunner As TaskRunner)
            _credentialManager = credentialManager
            _xmlManager = xmlManager
            _toolManager = toolManager
            _logger = logger
            _taskRunner = taskRunner
            _errorNotifier = New ErrorNotifier(xmlManager, logger)
            _dependencyManager = New TaskDependencyManager(logger, xmlManager, Me)
            _powerShellPathManager = ApplicationContainer.GetService(Of PowerShellPathManager)()

            ' Initialize mutex cleanup timer (run every 5 minutes)
            _mutexCleanupTimer = New Threading.Timer(AddressOf CleanupAbandonedMutexes, Nothing,
                                         TimeSpan.FromMinutes(5), TimeSpan.FromMinutes(5))

            ' Cleanup any abandoned mutexes from previous application runs
            CleanupAllAbandonedMutexes()

            _logger.LogInfo($"CustomScheduler initialized with enhanced lock management")
        End Sub

        ''' Cleanup all abandoned locks on application startup
        Private Sub CleanupAllAbandonedMutexes()
            Try
                _logger.LogInfo("Starting cleanup of abandoned locks from previous runs")

                Dim cleanedCount = 0
                Dim taskNames = _xmlManager.GetAllTaskNames()

                For Each taskName In taskNames
                    If TryCleanupAbandonedLock(taskName, forceRelease:=True) Then
                        cleanedCount += 1
                    End If
                Next

                If cleanedCount > 0 Then
                    _logger.LogWarning($"Cleaned up {cleanedCount} abandoned locks from previous runs")
                Else
                    _logger.LogInfo("No abandoned locks found during startup")
                End If

            Catch ex As Exception
                _logger.LogError($"Error during startup lock cleanup: {ex.Message}")
            End Try
        End Sub

        ''' Attempts to cleanup a specific abandoned lock.
        ''' forceRelease should only be True at startup or periodic cleanup (stale > 15 min),
        ''' NOT during the retry loop in ExecuteTaskAsync, to avoid breaking exclusion
        ''' for a legitimately running task.
        Private Function TryCleanupAbandonedLock(taskName As String, Optional forceRelease As Boolean = False) As Boolean
            Dim semaphoreName = _mutexBasePath & taskName
            Dim semaphore As Semaphore = Nothing

            Try
                semaphore = New Semaphore(1, 1, semaphoreName)

                ' Try to acquire with very short timeout to check if it's abandoned
                If semaphore.WaitOne(100) Then
                    ' Got the semaphore immediately - it wasn't actually in use
                    semaphore.Release()
                    Return False
                Else
                    If Not forceRelease Then
                        ' Without force-release, retry with a slightly longer timeout
                        ' in case the holder is in the process of releasing
                        semaphore?.Dispose()
                        semaphore = Nothing

                        semaphore = New Semaphore(1, 1, semaphoreName)
                        If semaphore.WaitOne(500) Then
                            semaphore.Release()
                            _logger.LogWarning($"Successfully cleaned up lock for task '{taskName}'")
                            Return True
                        End If

                        ' Still held - don't force-release, task is legitimately running
                        Return False
                    End If

                    ' Force-release path: only used at startup or for stale tasks.
                    ' Verify the task is not legitimately running in our process first.
                    Dim isRunningInProcess As Boolean = False
                    _taskRunning.TryGetValue(taskName, isRunningInProcess)

                    If isRunningInProcess Then
                        ' Task is actively running in our process - don't force-release
                        Return False
                    End If

                    ' Task is NOT running in our process but semaphore is held.
                    ' This means it was abandoned by a crashed process or previous session.
                    Try
                        semaphore.Release()
                        _logger.LogWarning($"Force-released abandoned lock for task '{taskName}'")
                        Return True
                    Catch ex As SemaphoreFullException
                        ' Semaphore is at max count - not actually held
                        Return False
                    End Try
                End If

            Catch ex As Exception
                _logger.LogError($"Error cleaning lock for task '{taskName}': {ex.Message}")
                Return False
            Finally
                semaphore?.Dispose()
            End Try

            Return False
        End Function

        ''' Timer callback to periodically cleanup abandoned mutexes
        Private Sub CleanupAbandonedMutexes(state As Object)
            Try
                Dim now = DateTime.Now
                Dim abandonedLocks = New List(Of String)()

                ' Check for locks that have been active too long
                For Each kvp In _activeMutexes
                    If (now - kvp.Value).TotalMinutes > STALE_TIMEOUT_MINUTES Then
                        abandonedLocks.Add(kvp.Key)
                    End If
                Next

                ' Attempt to cleanup abandoned locks
                For Each taskName In abandonedLocks
                    _logger.LogWarning($"Attempting to cleanup potentially abandoned lock for task: {taskName}")

                    If TryCleanupAbandonedLock(taskName, forceRelease:=True) Then
                        _activeMutexes.TryRemove(taskName, Nothing)
                        ' Also reset the task state
                        If _taskStates.TryGetValue(taskName, Nothing) Then
                            Dim taskState = _taskStates(taskName)
                            taskState.IsRunning = False
                            taskState.LastError = "Lock cleanup - task may have crashed"
                        End If
                    End If
                Next

            Catch ex As Exception
                _logger.LogError($"Error during periodic lock cleanup: {ex.Message}")
            End Try
        End Sub

        Public Sub AddTask(task As ScheduledTask)
            Try
                If task Is Nothing Then
                    Throw New ArgumentNullException(NameOf(task))
                End If

                _logger.LogInfo($"Adding new task: {task.Name}")

                ' Validate task
                If String.IsNullOrWhiteSpace(task.Name) Then
                    Throw New ArgumentException("Task name is required")
                End If

                If _tasks.ContainsKey(task.Name) Then
                    Throw New ArgumentException($"Task with name '{task.Name}' already exists")
                End If

                ' Ensure collections are initialized
                If task.Actions Is Nothing Then
                    task.Actions = New List(Of TaskAction)()
                End If
                If task.DailyTimes Is Nothing Then
                    task.DailyTimes = New List(Of TimeSpan)()
                End If
                If task.Parameters Is Nothing Then
                    task.Parameters = New Hashtable()
                End If

                ' Calculate initial next run time
                task.NextRunTime = task.CalculateNextRunTime()

                ' Add task to collection
                _tasks(task.Name) = task

                ' Save to XML
                _xmlManager.SaveTask(task)

                _logger.LogInfo($"Task {task.Name} added successfully")
            Catch ex As Exception
                _logger.LogError($"Error adding task: {ex.Message}")
                Throw New Exception("Failed to add task", ex)
            End Try
        End Sub

        Public Sub ClearTaskStates()
            _taskStates.Clear()
            _activeMutexes.Clear()
            _staleTaskAlerts.Clear()
            _logger.LogInfo("Task states, mutex tracking, and stale alert tracking cleared")
        End Sub

        Public Sub CheckAndExecuteTasks()
            Try
                Dim now = DateTime.Now

                ' Only run stale cleanup every few minutes instead of every tick
                If (now - _lastStaleCleanupTime).TotalSeconds >= STALE_CLEANUP_INTERVAL_SECONDS Then
                    _lastStaleCleanupTime = now
                    CleanupStaleTasks()
                End If

                ' Periodic memory usage monitoring
                CheckMemoryUsage(now)

                For Each task In _tasks.Values
                    Try
                        If task.NextRunTime <= now Then
                            Dim taskState = _taskStates.GetOrAdd(task.Name, New TaskState())

                            If Not taskState.IsRunning Then
                                _logger.LogInfo($"Task due for execution: {task.Name}")
                                ExecuteTaskAsync(task).ConfigureAwait(False)
                            End If
                        End If
                    Catch ex As Exception
                        _logger.LogError($"Error checking task {task.Name}: {ex.Message}")
                    End Try
                Next
            Catch ex As Exception
                _logger.LogError($"Error in CheckAndExecuteTasks: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Lightweight memory monitoring that logs usage every 5 minutes and sends
        ''' email alerts when thresholds are exceeded. Designed to detect progressive
        ''' memory leaks before they crash the server.
        ''' </summary>
        Private Sub CheckMemoryUsage(now As DateTime)
            Try
                If (now - _lastMemoryCheckTime).TotalSeconds < MEMORY_CHECK_INTERVAL_SECONDS Then
                    Return
                End If
                _lastMemoryCheckTime = now

                Using currentProcess = Diagnostics.Process.GetCurrentProcess()
                    Dim workingSetMB = currentProcess.WorkingSet64 \ (1024L * 1024L)
                    Dim privateMemoryMB = currentProcess.PrivateMemorySize64 \ (1024L * 1024L)

                    _logger.LogInfo($"[MemoryMonitor] Working Set: {workingSetMB} MB, Private Bytes: {privateMemoryMB} MB, GC Total: {GC.GetTotalMemory(False) \ (1024L * 1024L)} MB")

                    ' Only send alerts with a cooldown to avoid flooding
                    If (now - _lastMemoryAlertTime).TotalMinutes < MEMORY_ALERT_COOLDOWN_MINUTES Then
                        Return
                    End If

                    Dim level As String = Nothing
                    If privateMemoryMB > MEMORY_CRITICAL_THRESHOLD_MB Then
                        level = "CRITICAL"
                        _logger.LogError($"[MemoryMonitor] CRITICAL: Private Bytes {privateMemoryMB} MB exceeds {MEMORY_CRITICAL_THRESHOLD_MB} MB threshold")
                    ElseIf privateMemoryMB > MEMORY_WARNING_THRESHOLD_MB Then
                        level = "WARNING"
                        _logger.LogWarning($"[MemoryMonitor] WARNING: Private Bytes {privateMemoryMB} MB exceeds {MEMORY_WARNING_THRESHOLD_MB} MB threshold")
                    End If

                    If level IsNot Nothing Then
                        _lastMemoryAlertTime = now
                        Try
                            Dim notificationManager = ApplicationContainer.GetService(Of NotificationManager)()
                            If notificationManager IsNot Nothing Then
                                Dim body As New StringBuilder()
                                body.AppendLine($"LiteTask memory usage has exceeded the {level} threshold:")
                                body.AppendLine()
                                body.AppendLine($"  Working Set:  {workingSetMB} MB")
                                body.AppendLine($"  Private Bytes: {privateMemoryMB} MB")
                                body.AppendLine($"  GC Heap:       {GC.GetTotalMemory(False) \ (1024L * 1024L)} MB")
                                body.AppendLine($"  Threshold:     {If(level = "CRITICAL", MEMORY_CRITICAL_THRESHOLD_MB, MEMORY_WARNING_THRESHOLD_MB)} MB")
                                body.AppendLine()
                                body.AppendLine($"  Time: {now:yyyy-MM-dd HH:mm:ss}")
                                Dim runningTaskCount = _taskStates.Values.Where(Function(ts) ts.IsRunning).Count()
                                body.AppendLine($"  Running tasks: {runningTaskCount}")
                                body.AppendLine()
                                body.AppendLine("Action recommended: Investigate running tasks and consider restarting the service.")

                                notificationManager.QueueNotification(
                                    $"LiteTask {level}: Memory usage at {privateMemoryMB} MB",
                                    body.ToString(),
                                    If(level = "CRITICAL",
                                       NotificationManager.NotificationPriority.High,
                                       NotificationManager.NotificationPriority.Normal))
                            End If
                        Catch notifyEx As Exception
                            _logger.LogError($"[MemoryMonitor] Failed to send memory alert: {notifyEx.Message}")
                        End Try
                    End If
                End Using
            Catch ex As Exception
                _logger.LogError($"[MemoryMonitor] Error checking memory: {ex.Message}")
            End Try
        End Sub

        Private Sub CleanupStaleTasks()
            Try
                Dim staleTimeout = TimeSpan.FromMinutes(STALE_TIMEOUT_MINUTES)
                Dim now = DateTime.Now

                For Each taskEntry In _taskStates
                    If taskEntry.Value.IsRunning AndAlso
                   (now - taskEntry.Value.LastStartTime) > staleTimeout Then
                        _logger.LogWarning($"Cleaning up stale task state: {taskEntry.Key}")

                        ' Send email alert if not already sent for this stale instance
                        SendStaleTaskAlert(taskEntry.Key, taskEntry.Value, staleTimeout)

                        taskEntry.Value.IsRunning = False
                        taskEntry.Value.LastError = "Task cleaned up due to timeout"

                        ' Also remove from mutex tracking
                        _activeMutexes.TryRemove(taskEntry.Key, Nothing)
                    End If
                Next
            Catch ex As Exception
                _logger.LogError($"Error cleaning up stale tasks: {ex.Message}")
            End Try
        End Sub

        ''' Sends an email alert for a stale task, with duplicate prevention
        Private Sub SendStaleTaskAlert(taskName As String, taskState As TaskState, staleTimeout As TimeSpan)
            Try
                Dim lastAlertTime As DateTime = Nothing
                Dim now = DateTime.Now

                ' Check if we already sent an alert for this task recently (within the last hour)
                If _staleTaskAlerts.TryGetValue(taskName, lastAlertTime) Then
                    If (now - lastAlertTime).TotalHours < 1 Then
                        ' Already sent an alert recently, don't spam
                        _logger.LogInfo($"Skipping duplicate stale task alert for {taskName} (last alert: {lastAlertTime})")
                        Return
                    End If
                End If

                ' Get notification manager
                Dim notificationManager = ApplicationContainer.GetService(Of NotificationManager)()
                If notificationManager Is Nothing Then
                    _logger.LogWarning("NotificationManager not available for stale task alert")
                    Return
                End If

                ' Build alert message
                Dim subject = $"ALERT: Stale Task Detected - {taskName}"
                Dim body As New StringBuilder()
                body.AppendLine($"A stale task has been detected and cleaned up:")
                body.AppendLine()
                body.AppendLine($"Task Name: {taskName}")
                body.AppendLine($"Started At: {taskState.LastStartTime:yyyy-MM-dd HH:mm:ss}")
                body.AppendLine($"Detection Time: {now:yyyy-MM-dd HH:mm:ss}")
                body.AppendLine($"Running Duration: {(now - taskState.LastStartTime).TotalMinutes:F1} minutes")
                body.AppendLine($"Timeout Threshold: {staleTimeout.TotalMinutes} minutes")
                body.AppendLine()
                body.AppendLine($"Status: {taskState.StatusMessage}")
                If Not String.IsNullOrEmpty(taskState.LastError) Then
                    body.AppendLine($"Last Error: {taskState.LastError}")
                End If
                body.AppendLine()
                body.AppendLine("Action Taken: Task state has been reset and mutex released.")
                body.AppendLine()
                body.AppendLine("This typically indicates:")
                body.AppendLine("  - The task process may be hung or stuck")
                body.AppendLine("  - An executable may still be running in the background")
                body.AppendLine("  - The task may have crashed without cleaning up properly")
                body.AppendLine()
                body.AppendLine("Recommended Actions:")
                body.AppendLine("  1. Check task manager for hung processes")
                body.AppendLine("  2. Review task logs for errors")
                body.AppendLine("  3. Verify task executable and parameters")
                body.AppendLine("  4. Consider increasing timeout if task legitimately needs longer")

                ' Queue the notification
                notificationManager.QueueNotification(subject, body.ToString(), NotificationManager.NotificationPriority.High)

                ' Update alert tracking
                _staleTaskAlerts(taskName) = now

                _logger.LogInfo($"Stale task alert sent for {taskName}")

            Catch ex As Exception
                _logger.LogError($"Error sending stale task alert for {taskName}: {ex.Message}")
            End Try
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not _disposed Then
                If disposing Then
                    ' Stop the cleanup timer
                    _mutexCleanupTimer?.Dispose()
                    ' Clean up managed resources
                    For Each task In _tasks.Values
                        If TypeOf task Is IDisposable Then
                            DirectCast(task, IDisposable).Dispose()
                        End If
                    Next
                    _tasks.Clear()
                    ' Cleanup any remaining locks
                    For Each taskName In _activeMutexes.Keys
                        TryCleanupAbandonedLock(taskName, forceRelease:=True)
                    Next
                    _activeMutexes.Clear()

                    ' Dispose all SemaphoreSlim objects in _taskLocks
                    For Each kvp In _taskLocks
                        If TypeOf kvp.Value Is SemaphoreSlim Then
                            DirectCast(kvp.Value, SemaphoreSlim).Dispose()
                        End If
                    Next
                    _taskLocks.Clear()

                    ' Clear remaining tracking dictionaries
                    _taskStates.Clear()
                    _taskRunning.Clear()
                    _staleTaskAlerts.Clear()

                    ' Release the processing semaphores
                    _schedulerLock.Dispose()
                    _processingLock.Dispose()
                End If

                ' Clean up unmanaged resources
                _disposed = True
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub


        Public Async Function ExecuteTask(task As ScheduledTask) As Task
            Dim taskLock = _taskLocks.GetOrAdd(task.Name, New SemaphoreSlim(1, 1))

            Try
                Await taskLock.WaitAsync()

                task = GetTask(task.Name) ' Get fresh copy
                If task Is Nothing Then Return

                _logger.LogInfo($"Executing task: {task.Name}")

                For Each action In task.Actions
                    Try
                        Dim credential = If(Not String.IsNullOrEmpty(task.CredentialTarget),
                                     _credentialManager.GetCredential(task.CredentialTarget, task.AccountType),
                                     Nothing)

                        Dim success = False
                        Select Case action.Type
                            Case TaskType.SQL
                                success = Await _taskRunner.ExecuteSqlTask(action, credential)
                            Case TaskType.PowerShell
                                success = Await _taskRunner.ExecutePowerShellTask(action, credential)
                            Case TaskType.Batch
                                success = Await _taskRunner.ExecuteBatchTask(action, credential)
                            Case TaskType.Executable
                                success = Await _taskRunner.ExecuteExecutableTask(action, credential)
                            Case Else
                                Throw New NotSupportedException($"Task type {action.Type} is not supported")
                        End Select

                        If Not success AndAlso Not action.ContinueOnError Then
                            Throw New Exception($"Action failed: {action.Name}")
                        End If
                    Catch ex As Exception
                        _logger.LogError($"Action {action.Name} failed: {ex.Message}")
                        If Not action.ContinueOnError Then Throw
                    End Try
                Next

                task.NextRunTime = task.CalculateNextRunTime()
                UpdateTask(task)
                SaveTasks()

            Finally
                taskLock.Release()
            End Try
        End Function

        ''' Enhanced internal execution method with improved mutex handling
        ''' Returns True if the task actually executed, False if it was skipped (e.g. lock not acquired).
        Public Async Function ExecuteTaskAsync(_task As ScheduledTask) As Task(Of Boolean)
            Dim semaphoreName = _mutexBasePath & _task.Name
            Dim semaphore As Semaphore = Nothing
            Dim hasLock As Boolean = False
            Dim retryCount As Integer = 0
            Dim taskState = _taskStates.GetOrAdd(_task.Name, New TaskState())

            Try
                ' Record acquisition attempt
                _activeMutexes.TryAdd(_task.Name, DateTime.Now)

                ' Try to acquire semaphore with retries for abandoned scenarios
                ' Using Semaphore instead of Mutex because Mutex has thread affinity:
                ' in async code, the thread that releases may differ from the one that acquired.
                Do While retryCount < MAX_CLEANUP_RETRIES
                    Try
                        semaphore = New Semaphore(1, 1, semaphoreName)
                        hasLock = semaphore.WaitOne(TimeSpan.FromSeconds(MUTEX_TIMEOUT_SECONDS))

                        If hasLock Then
                            Exit Do ' Successfully acquired
                        Else
                            _logger.LogWarning($"Task {_task.Name} lock timeout (attempt {retryCount + 1}). Trying cleanup...")
                            semaphore?.Dispose()
                            semaphore = Nothing

                            ' Attempt to cleanup potentially abandoned lock
                            If TryCleanupAbandonedLock(_task.Name) Then
                                _logger.LogInfo($"Cleaned up abandoned lock for {_task.Name}, retrying...")
                                retryCount += 1
                                Await Task.Delay(1000) ' Brief delay before retry
                            Else
                                ' Genuine busy condition - task is actually running
                                _logger.LogInfo($"Task {_task.Name} is legitimately running. Will retry later.")
                                Exit Do
                            End If
                        End If

                    Catch ex As Exception
                        _logger.LogError($"Error acquiring lock for task {_task.Name}: {ex.Message}")
                        semaphore?.Dispose()
                        semaphore = Nothing
                        Exit Do
                    End Try
                Loop

                If Not hasLock Then
                    _logger.LogInfo($"Could not acquire lock for task {_task.Name} after {retryCount} retries. Will retry later.")
                    _activeMutexes.TryRemove(_task.Name, Nothing)
                    Return False
                End If

                ' Update task state and tracking
                taskState.IsRunning = True
                taskState.LastStartTime = DateTime.Now
                taskState.StatusMessage = "Starting execution"
                _activeMutexes(_task.Name) = DateTime.Now
                _logger.LogInfo($"Lock acquired for task: {_task.Name}")
                If _taskRunning.TryAdd(_task.Name, True) Then
                    Try
                        ' Get credential if needed
                        Dim credential As CredentialInfo = Nothing
                        If Not String.IsNullOrEmpty(_task.CredentialTarget) AndAlso
                   Not String.IsNullOrEmpty(_task.AccountType) Then
                            credential = _credentialManager.GetCredential(_task.CredentialTarget, _task.AccountType)
                        End If

                        For Each action In _task.Actions
                            action.Status = TaskActionStatus.Running
                            action.LastRunTime = DateTime.Now

                            Try
                                Dim success = Await ExecuteTaskAction(action, credential)
                                If success Then
                                    action.Status = TaskActionStatus.Completed
                                    _logger.LogInfo($"Action {action.Name} completed successfully")
                                Else
                                    Throw New Exception($"Action failed: {action.Name}")
                                End If
                            Catch ex As Exception
                                action.Status = TaskActionStatus.Failed
                                _logger.LogError($"Action {action.Name} failed: {ex.Message}")

                                Try
                                    Dim notificationManager = ApplicationContainer.GetService(Of NotificationManager)()
                                    notificationManager?.QueueNotification(
                                $"Task Action Failed: {_task.Name} - {action.Name}",
                                $"Error: {ex.Message}",
                                NotificationManager.NotificationPriority.High)
                                Catch notifyEx As Exception
                                    _logger.LogError($"Failed to send notification: {notifyEx.Message}")
                                End Try

                                If Not action.ContinueOnError Then
                                    Throw
                                End If
                            End Try
                        Next

                        _task.UpdateNextRunTime()
                        SaveTasks()
                        taskState.StatusMessage = "Execution completed successfully"

                        ' Make a copy of the task for the event to avoid threading issues
                        Dim taskCopy = _task.Clone()

                        ' Raise event safely
                        Try
                            Dim handler = TaskCompletedEvent
                            If handler IsNot Nothing Then
                                If System.Windows.Forms.Application.MessageLoop Then
                                    ' We're in a Windows Forms context, use Control.BeginInvoke
                                    If TypeOf handler.Target Is Control Then
                                        DirectCast(handler.Target, Control).BeginInvoke(Sub()
                                                                                            handler(Me, New TaskCompletedEventArgs(taskCopy))
                                                                                        End Sub)
                                    Else
                                        ' Use SynchronizationContext if available
                                        Dim context = SynchronizationContext.Current
                                        If context IsNot Nothing Then
                                            context.Post(Sub() handler(Me, New TaskCompletedEventArgs(taskCopy)), Nothing)
                                        Else
                                            ' Fallback to direct invocation
                                            handler(Me, New TaskCompletedEventArgs(taskCopy))
                                        End If
                                    End If
                                Else
                                    ' Not in UI context, invoke directly
                                    handler(Me, New TaskCompletedEventArgs(taskCopy))
                                End If
                            End If
                        Catch ex As Exception
                            _logger.LogError($"Error raising TaskCompleted event: {ex.Message}")
                        End Try

                    Finally
                        _taskRunning.TryRemove(_task.Name, False)
                        taskState.LastEndTime = DateTime.Now
                        taskState.IsRunning = False
                        ' Clear any stale task alert tracking when task completes normally
                        _staleTaskAlerts.TryRemove(_task.Name, Nothing)
                    End Try

                    Return True
                Else
                    ' Task is already tracked as running in-process (secondary guard)
                    _logger.LogWarning($"Task {_task.Name} is already running in-process, skipping.")
                    Return False
                End If

            Catch ex As Exception
                _logger.LogError($"Fatal error in ExecuteTaskAsync for {_task.Name}: {ex.Message}")
                taskState.IsRunning = False
                taskState.LastError = ex.Message

                ' CRITICAL: Always advance the task schedule on failure to prevent
                ' the task from re-executing every 60 seconds and flooding emails.
                ' Without this, NextRunTime stays in the past and CheckAndExecuteTasks
                ' will keep triggering the same failing task indefinitely.
                Try
                    _task.UpdateNextRunTime()
                    SaveTasks()
                Catch schedEx As Exception
                    _logger.LogError($"Failed to update next run time for {_task.Name}: {schedEx.Message}")
                End Try

                ' Send email notification for fatal errors
                Try
                    Dim notificationManager = ApplicationContainer.GetService(Of NotificationManager)()
                    If notificationManager IsNot Nothing Then
                        Dim errorBody As New StringBuilder()
                        errorBody.AppendLine($"A fatal error occurred during task execution:")
                        errorBody.AppendLine()
                        errorBody.AppendLine($"Task Name: {_task.Name}")
                        errorBody.AppendLine($"Error Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
                        errorBody.AppendLine($"Error Message: {ex.Message}")
                        errorBody.AppendLine()
                        errorBody.AppendLine($"Stack Trace:")
                        errorBody.AppendLine(ex.StackTrace)

                        notificationManager.QueueNotification(
                            $"FATAL ERROR: Task Execution Failed - {_task.Name}",
                            errorBody.ToString(),
                            NotificationManager.NotificationPriority.High)
                    End If
                Catch notifyEx As Exception
                    _logger.LogError($"Failed to send fatal error notification: {notifyEx.Message}")
                End Try

                Throw
            Finally
                ' Cleanup in Finally block
                Try
                    If hasLock AndAlso semaphore IsNot Nothing Then
                        Try
                            semaphore.Release()
                            _logger.LogInfo($"Released lock for task: {_task.Name}")
                        Catch semEx As SemaphoreFullException
                            ' Semaphore was already released
                            _logger.LogWarning($"Lock already released for task: {_task.Name}")
                        Catch semEx As ObjectDisposedException
                            ' Semaphore was already disposed
                            _logger.LogWarning($"Lock already disposed for task: {_task.Name}")
                        End Try
                    End If
                Catch ex As Exception
                    _logger.LogError($"Error during lock cleanup for {_task.Name}: {ex.Message}")
                Finally
                    ' Dispose semaphore safely
                    Try
                        semaphore?.Dispose()
                    Catch disposeEx As Exception
                        _logger.LogWarning($"Error disposing lock for {_task.Name}: {disposeEx.Message}")
                    End Try

                    _activeMutexes.TryRemove(_task.Name, Nothing)
                    If taskState.IsRunning Then
                        taskState.IsRunning = False
                        taskState.StatusMessage = "Completed with cleanup"
                    End If
                End Try
            End Try
        End Function

        ' Helper method for executing individual actions
        Private Async Function ExecuteTaskAction(action As TaskAction, credential As CredentialInfo) As Task(Of Boolean)
            Return Await Task.Run(Async Function()
                                      Try
                                          Select Case action.Type
                                              Case TaskType.PowerShell
                                                  Return Await _taskRunner.ExecutePowerShellTask(action, credential)
                                              Case TaskType.Batch
                                                  Return Await _taskRunner.ExecuteBatchTask(action, credential)
                                              Case TaskType.SQL
                                                  Return Await _taskRunner.ExecuteSqlTask(action, credential)
                                              Case TaskType.Executable
                                                  Return Await _taskRunner.ExecuteExecutableTask(action, credential)
                                              Case Else
                                                  Throw New NotSupportedException($"Task type {action.Type} is not supported")
                                          End Select
                                      Catch ex As Exception
                                          _logger.LogError($"Error executing action {action.Name}: {ex.Message}")
                                          Throw
                                      End Try
                                  End Function)
        End Function

        Public Function GetAllTasks() As IEnumerable(Of ScheduledTask)
            Try
                _logger.LogInfo("Getting all tasks")

                ' First load all tasks from XML to ensure memory cache is up to date
                Dim taskNames = _xmlManager.GetAllTaskNames()
                For Each taskName In taskNames
                    If Not _tasks.ContainsKey(taskName) Then
                        Dim task = _xmlManager.LoadTask(taskName)
                        If task IsNot Nothing Then
                            _tasks(taskName) = task
                        End If
                    End If
                Next

                ' Return tasks from memory
                Return _tasks.Values.Where(Function(t) t IsNot Nothing).ToList()

            Catch ex As Exception
                _logger.LogError($"Error getting all tasks: {ex.Message}")
                _logger.LogError($"StackTrace: {ex.StackTrace}")
                Return New List(Of ScheduledTask)()
            End Try
        End Function

        Public Function GetTask(taskName As String) As ScheduledTask
            Try
                _logger.LogInfo($"Getting task: {taskName}")

                ' First try to get from memory
                Dim task As ScheduledTask = Nothing
                If _tasks.TryGetValue(taskName, task) Then
                    Return task
                End If

                ' If not in memory, try to load from XML
                task = _xmlManager.LoadTask(taskName)
                If task IsNot Nothing Then
                    ' Cache the task in memory
                    _tasks(taskName) = task
                    Return task
                End If

                _logger.LogWarning($"Task not found: {taskName}")
                Return Nothing

            Catch ex As Exception
                _logger.LogError($"Error getting task {taskName}: {ex.Message}")
                _logger.LogError($"StackTrace: {ex.StackTrace}")
                Return Nothing
            End Try
        End Function

        Public Sub LoadTasks()
            Try
                _logger.LogInfo("Loading tasks from XML file")
                Dim taskNames = _xmlManager.GetAllTaskNames()
                If taskNames.Count = 0 Then
                    _logger.LogInfo("No tasks found in the XML file")
                    Return
                End If

                For Each taskName In taskNames
                    Dim task = _xmlManager.LoadTask(taskName)
                    If task IsNot Nothing Then
                        _tasks(taskName) = task
                    End If
                Next
                _logger.LogInfo($"Successfully loaded {_tasks.Count} tasks")
            Catch ex As Exception
                _logger.LogError($"Error loading tasks: {ex.Message}")
                _logger.LogError($"StackTrace: {ex.StackTrace}")
            End Try
        End Sub

        Private Sub HandleTaskError(task As ScheduledTask, ex As Exception)
            _logger.LogError($"Error executing task {task.Name}: {ex.Message}")
            _logger.LogError($"Stack trace: {ex.StackTrace}")

            Try
                Dim notificationManager = ApplicationContainer.GetService(Of NotificationManager)()
                notificationManager?.QueueNotification(
           $"Task Execution Error: {task.Name}",
           $"Error: {ex.Message}",
           NotificationManager.NotificationPriority.High)
            Catch notifyEx As Exception
                _logger.LogError($"Failed to send error notification: {notifyEx.Message}")
            End Try
        End Sub

        Public Sub RemoveTask(taskName As String)
            Try
                _logger?.LogInfo($"Attempting to remove task: {taskName}")

                If String.IsNullOrEmpty(taskName) Then
                    Throw New ArgumentException("Task name cannot be null or empty", NameOf(taskName))
                End If

                ' Try to remove from memory
                Dim task As ScheduledTask = Nothing
                If _tasks.TryRemove(taskName, task) Then
                    _logger?.LogInfo($"Task {taskName} removed from scheduler cache")
                Else
                    _logger?.LogWarning($"Task {taskName} not found in scheduler cache")
                End If

                ' Clean up all associated dictionary entries to prevent memory accumulation
                _taskStates.TryRemove(taskName, Nothing)
                _taskRunning.TryRemove(taskName, Nothing)
                _staleTaskAlerts.TryRemove(taskName, Nothing)
                _activeMutexes.TryRemove(taskName, Nothing)

                Dim removedLock As Object = Nothing
                If _taskLocks.TryRemove(taskName, removedLock) Then
                    ' Dispose the SemaphoreSlim if it was used as the lock object
                    If TypeOf removedLock Is SemaphoreSlim Then
                        DirectCast(removedLock, SemaphoreSlim).Dispose()
                    End If
                End If

                ' Try to remove from XML storage
                Try
                    _xmlManager.DeleteTask(taskName)
                    _logger?.LogInfo($"Task {taskName} removed from XML storage")
                Catch xmlEx As Exception
                    _logger?.LogError($"Error removing task {taskName} from XML storage: {xmlEx.Message}")
                    Throw
                End Try

            Catch ex As Exception
                _logger?.LogError($"Error removing task {taskName}: {ex.Message}")
                _logger?.LogError($"StackTrace: {ex.StackTrace}")
                Throw
            End Try
        End Sub

        ' Primary public entry point for running tasks
        Public Async Function RunTaskAsync(task As ScheduledTask) As Task
            If task Is Nothing Then
                _logger.LogError("Cannot run null task")
                Return
            End If

            Try
                _logger.LogInfo($"Starting execution of task: {task.Name}")
                Dim executed = Await ExecuteTaskAsync(task)

                If Not executed Then
                    _logger.LogWarning($"Task {task.Name} was not executed (already running or lock unavailable)")
                    Return
                End If

                ' Handle chained tasks
                If task.NextTaskId.HasValue AndAlso task.NextTaskId.Value > 0 Then
                    Dim nextTask = GetTask(task.NextTaskId.Value.ToString())
                    If nextTask IsNot Nothing Then
                        Await ExecuteTaskAsync(nextTask)
                    End If
                End If

                ' Send success notification if needed
                Try
                    Dim notificationManager = ApplicationContainer.GetService(Of NotificationManager)()
                    notificationManager?.QueueNotification(
                        $"Task Completed: {task.Name}",
                        $"Task executed successfully at {DateTime.Now}",
                        NotificationManager.NotificationPriority.Normal)
                Catch notifyEx As Exception
                    _logger.LogError($"Failed to send success notification: {notifyEx.Message}")
                End Try

            Catch ex As Exception
                _logger.LogError($"Error running task {task.Name}: {ex.Message}")
                _logger.LogError($"Stack trace: {ex.StackTrace}")

                Try
                    Dim notificationManager = ApplicationContainer.GetService(Of NotificationManager)()
                    If notificationManager IsNot Nothing Then
                        notificationManager.QueueNotification(
                            $"Task Failed: {task.Name}",
                            $"Error: {ex.Message}{Environment.NewLine}Stack Trace: {ex.StackTrace}",
                            NotificationManager.NotificationPriority.High)
                    End If
                Catch notifyEx As Exception
                    _logger.LogError($"Failed to send error notification: {notifyEx.Message}")
                End Try

                Throw
            End Try
        End Function

        Public Sub SaveTasks()
            Try
                '_logger.LogInfo("Starting to save tasks")

                If _xmlManager Is Nothing Then
                    Throw New InvalidOperationException("XMLManager is not initialized")
                End If

                ' Get current tasks
                Dim currentTasks = GetAllTasks()

                ' First, load existing task names from XML
                Dim existingTaskNames = _xmlManager.GetAllTaskNames()

                ' Find tasks to delete (tasks in XML but not in current tasks)
                For Each existingName In existingTaskNames
                    If Not currentTasks.Any(Function(t) t.Name = existingName) Then
                        _logger.LogInfo($"Deleting obsolete task from XML: {existingName}")
                        _xmlManager.DeleteTask(existingName)
                    End If
                Next

                ' Save current tasks
                For Each task In currentTasks
                    If task IsNot Nothing Then
                        _xmlManager.SaveTask(task)
                        '_logger.LogInfo($"Saved task: {task.Name}")
                    End If
                Next

                _logger.LogInfo("Tasks saved successfully")
            Catch ex As Exception
                _logger.LogError($"Error saving tasks: {ex.Message}")
                _logger.LogError($"StackTrace: {ex.StackTrace}")
                Throw New Exception("Failed to save tasks", ex)
            End Try
        End Sub

        Public Sub UpdateTask(updatedTask As ScheduledTask)
            Try
                If updatedTask Is Nothing Then
                    Throw New ArgumentNullException(NameOf(updatedTask))
                End If

                _logger.LogInfo($"Updating task: {updatedTask.Name}")

                If Not _tasks.ContainsKey(updatedTask.Name) Then
                    Throw New ArgumentException($"Task '{updatedTask.Name}' not found")
                End If

                ' Update next run time
                updatedTask.NextRunTime = updatedTask.CalculateNextRunTime()

                ' Update task in collection
                _tasks(updatedTask.Name) = updatedTask

                ' Save to XML
                _xmlManager.SaveTask(updatedTask)

                _logger.LogInfo($"Task {updatedTask.Name} updated successfully")
            Catch ex As Exception
                _logger.LogError($"Error updating task: {ex.Message}")
                Throw New Exception("Failed to update task", ex)
            End Try
        End Sub

    End Class
End Namespace