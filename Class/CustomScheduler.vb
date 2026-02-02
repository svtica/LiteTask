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

            _logger.LogInfo($"CustomScheduler initialized with enhanced mutex management")
        End Sub

        ''' Cleanup all abandoned mutexes on application startup
        Private Sub CleanupAllAbandonedMutexes()
            Try
                _logger.LogInfo("Starting cleanup of abandoned mutexes from previous runs")
                
                Dim cleanedCount = 0
                Dim taskNames = _xmlManager.GetAllTaskNames()
                
                For Each taskName In taskNames
                    If TryCleanupAbandonedMutex(taskName) Then
                        cleanedCount += 1
                    End If
                Next
                
                If cleanedCount > 0 Then
                    _logger.LogWarning($"Cleaned up {cleanedCount} abandoned mutexes from previous runs")
                Else
                    _logger.LogInfo("No abandoned mutexes found during startup")
                End If
                
            Catch ex As Exception
                _logger.LogError($"Error during startup mutex cleanup: {ex.Message}")
            End Try
        End Sub

        ''' Attempts to cleanup a specific abandoned mutex
        Private Function TryCleanupAbandonedMutex(taskName As String) As Boolean
            Dim mutexName = _mutexBasePath & taskName
            Dim mutex As Mutex = Nothing
            
            Try
                mutex = New Mutex(False, mutexName)
                
                ' Try to acquire mutex with very short timeout to check if it's abandoned
                If mutex.WaitOne(100) Then
                    ' Got the mutex immediately - it wasn't actually in use
                    mutex.ReleaseMutex()
                    Return False
                Else
                    ' Mutex appears to be held - try more aggressive cleanup
                    mutex.Dispose()
                    mutex = Nothing
                    
                    ' Create a new mutex to try to force cleanup
                    mutex = New Mutex(False, mutexName)
                    If mutex.WaitOne(500) Then
                        mutex.ReleaseMutex()
                        _logger.LogWarning($"Successfully cleaned abandoned mutex for task '{taskName}'")
                        Return True
                    End If
                End If
                
            Catch ex As AbandonedMutexException
                _logger.LogWarning($"Found and cleaned abandoned mutex for task '{taskName}'")
                Return True
            Catch ex As Exception
                _logger.LogError($"Error cleaning mutex for task '{taskName}': {ex.Message}")
                Return False
            Finally
                mutex?.Dispose()
            End Try
            
            Return False
        End Function

        ''' Timer callback to periodically cleanup abandoned mutexes
        Private Sub CleanupAbandonedMutexes(state As Object)
            Try
                Dim now = DateTime.Now
                Dim abandonedMutexes = New List(Of String)()
                
                ' Check for mutexes that have been active too long
                For Each kvp In _activeMutexes
                    If (now - kvp.Value).TotalMinutes > STALE_TIMEOUT_MINUTES Then
                        abandonedMutexes.Add(kvp.Key)
                    End If
                Next
                
                ' Attempt to cleanup abandoned mutexes
                For Each taskName In abandonedMutexes
                    _logger.LogWarning($"Attempting to cleanup potentially abandoned mutex for task: {taskName}")
                    
                    If TryCleanupAbandonedMutex(taskName) Then
                        _activeMutexes.TryRemove(taskName, Nothing)
                        ' Also reset the task state
                        If _taskStates.TryGetValue(taskName, Nothing) Then
                            Dim taskState = _taskStates(taskName)
                            taskState.IsRunning = False
                            taskState.LastError = "Mutex cleanup - task may have crashed"
                        End If
                    End If
                Next
                
            Catch ex As Exception
                _logger.LogError($"Error during periodic mutex cleanup: {ex.Message}")
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
                    ' Cleanup any remaining mutexes
                    For Each taskName In _activeMutexes.Keys
                        TryCleanupAbandonedMutex(taskName)
                    Next
                    _activeMutexes.Clear()
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
        Public Async Function ExecuteTaskAsync(_task As ScheduledTask) As Task
            Dim mutexName = _mutexBasePath & _task.Name
            Dim mutex As Mutex = Nothing
            Dim hasLock As Boolean = False
            Dim retryCount As Integer = 0
            Dim taskState = _taskStates.GetOrAdd(_task.Name, New TaskState())

            Try
                ' Record mutex acquisition attempt
                _activeMutexes.TryAdd(_task.Name, DateTime.Now)
                
                ' Try to acquire mutex with retries for abandoned mutex scenarios
                Do While retryCount < MAX_CLEANUP_RETRIES
                    Try
                        mutex = New Mutex(False, mutexName)
                        hasLock = mutex.WaitOne(TimeSpan.FromSeconds(MUTEX_TIMEOUT_SECONDS))
                        
                        If hasLock Then
                            Exit Do ' Successfully acquired mutex
                        Else
                            _logger.LogWarning($"Task {_task.Name} mutex timeout (attempt {retryCount + 1}). Trying cleanup...")
                            mutex?.Dispose()
                            mutex = Nothing
                            
                            ' Attempt to cleanup potentially abandoned mutex
                            If TryCleanupAbandonedMutex(_task.Name) Then
                                _logger.LogInfo($"Cleaned up abandoned mutex for {_task.Name}, retrying...")
                                retryCount += 1
                                Await Task.Delay(1000) ' Brief delay before retry
                            Else
                                ' Genuine busy condition - task is actually running
                                _logger.LogInfo($"Task {_task.Name} is legitimately running. Will retry later.")
                                Exit Do
                            End If
                        End If
                        
                    Catch ex As AbandonedMutexException
                        _logger.LogWarning($"Recovered from abandoned mutex for task: {_task.Name}")
                        hasLock = True
                        Exit Do
                    Catch ex As Exception
                        _logger.LogError($"Error acquiring mutex for task {_task.Name}: {ex.Message}")
                        mutex?.Dispose()
                        mutex = Nothing
                        Exit Do
                    End Try
                Loop

                If Not hasLock Then
                    _logger.LogInfo($"Could not acquire mutex for task {_task.Name} after {retryCount} retries. Will retry later.")
                    _activeMutexes.TryRemove(_task.Name, Nothing)
                    Return
                End If

                ' Update task state and mutex tracking
                taskState.IsRunning = True
                taskState.LastStartTime = DateTime.Now
                taskState.StatusMessage = "Starting execution"
                _activeMutexes(_task.Name) = DateTime.Now
                _logger.LogInfo($"Mutex acquired for task: {_task.Name}")
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
                End If

            Catch ex As Exception
                _logger.LogError($"Fatal error in ExecuteTaskAsync for {_task.Name}: {ex.Message}")
                taskState.IsRunning = False
                taskState.LastError = ex.Message

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
                ' Enhanced cleanup in Finally block
                Try
                    If hasLock AndAlso mutex IsNot Nothing Then
                        Try
                            mutex.ReleaseMutex()
                            _logger.LogInfo($"Released mutex for task: {_task.Name}")
                        Catch mutexEx As ApplicationException
                            ' Mutex was not owned by the current thread or already released
                            _logger.LogWarning($"Mutex already released or not owned for task: {_task.Name}")
                        Catch mutexEx As ObjectDisposedException
                            ' Mutex was already disposed
                            _logger.LogWarning($"Mutex already disposed for task: {_task.Name}")
                        End Try
                    End If
                Catch ex As Exception
                    _logger.LogError($"Error during mutex cleanup for {_task.Name}: {ex.Message}")
                Finally
                    ' Dispose mutex safely
                    Try
                        mutex?.Dispose()
                    Catch disposeEx As Exception
                        _logger.LogWarning($"Error disposing mutex for {_task.Name}: {disposeEx.Message}")
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
                Await ExecuteTaskAsync(task)

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