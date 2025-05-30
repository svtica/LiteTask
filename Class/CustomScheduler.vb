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
        Private Const CHECK_INTERVAL_SECONDS As Integer = 300
        Private ReadOnly _processingLock As New SemaphoreSlim(1, 1)
        Private _isProcessing As Boolean = False


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

            ' _logger?.LogInfo($"CustomScheduler initialized with execution tool: {_toolManager.CurrentExecutionTool}")
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
            _logger.LogInfo("Task states cleared")
        End Sub

        Public Sub CheckAndExecuteTasks()
            Try
                CleanupStaleTasks()
                Dim now = DateTime.Now

                For Each task In _tasks.Values
                    Try
                        If task.NextRunTime <= now Then
                            _logger.LogInfo($"Checking task for execution: {task.Name}")
                            Dim taskState = _taskStates.GetOrAdd(task.Name, New TaskState())

                            If Not taskState.IsRunning Then
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
                        taskEntry.Value.IsRunning = False
                        taskEntry.Value.LastError = "Task cleaned up due to timeout"
                    End If
                Next
            Catch ex As Exception
                _logger.LogError($"Error cleaning up stale tasks: {ex.Message}")
            End Try
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not _disposed Then
                If disposing Then
                    ' Clean up managed resources
                    For Each task In _tasks.Values
                        If TypeOf task Is IDisposable Then
                            DirectCast(task, IDisposable).Dispose()
                        End If
                    Next
                    _tasks.Clear()
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

        ' Internal execution method
        Public Async Function ExecuteTaskAsync(_task As ScheduledTask) As Task
            Dim mutexName = _mutexBasePath & _task.Name
            Dim mutex As Mutex = Nothing
            Dim hasLock As Boolean = False
            Dim retryDelay = TimeSpan.FromSeconds(30)
            Dim taskState = _taskStates.GetOrAdd(_task.Name, New TaskState())

            Try
                mutex = New Mutex(False, mutexName)
                hasLock = mutex.WaitOne(1000)

                If Not hasLock Then
                    _logger.LogInfo($"Task {_task.Name} is already running or has an abandoned mutex. Will retry in {retryDelay.TotalSeconds} seconds")
                    Await Task.Delay(retryDelay)
                    Return
                End If

                taskState.IsRunning = True
                taskState.LastStartTime = DateTime.Now
                taskState.StatusMessage = "Starting execution"

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
                    End Try
                End If

            Finally
                If hasLock Then
                    mutex?.ReleaseMutex()
                End If
                mutex?.Dispose()
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