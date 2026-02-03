Imports System.Drawing
Imports System.ServiceProcess

Namespace LiteTask
    Public Class MainForm
        Inherits Form

        Private ReadOnly _customScheduler As CustomScheduler
        Private ReadOnly _xmlManager As XMLManager
        Private ReadOnly _credentialManager As CredentialManager
        Private ReadOnly _toolManager As ToolManager
        Private ReadOnly _logger As Logger
        Private ReadOnly _powershellManager As RunTab
        Private ReadOnly _sqlManager As SqlTab
        Private ReadOnly _translationManager As TranslationManager
        Private _task As ScheduledTask
        Private ReadOnly _notificationManager As NotificationManager
        Private _autoRefreshTimer As System.Windows.Forms.Timer
        Private ReadOnly _statusTimer As System.Windows.Forms.Timer
        Private ReadOnly _schedulerTimer As System.Windows.Forms.Timer
        Private _recurrenceType As ScheduledTask.RecurrenceType
        Private _taskType As ScheduledTask.TaskType
        Private _taskTabPage As TabPage
        Private _defaultAccountType As String = "Current User"
        'Private _logTextBox As TextBox
        Private _isGuiMode As Boolean
        Private _contextMenuStrip As ContextMenuStrip
        Private _taskTableLayout As TableLayoutPanel
        Private _splitContainer As SplitContainer
        Private _taskPanel As Panel
        Private _buttonPanel As FlowLayoutPanel
        Private _tabPage As TabPage
        Private WithEvents _exitMenuItem As ToolStripMenuItem

        Private _refreshPending As Boolean = False
        Private _lastRefreshTime As DateTime = DateTime.MinValue
        Private Const MIN_REFRESH_INTERVAL_MS As Integer = 2000

        Private Const SERVICE_NAME As String = "LiteTaskService"
        Private _lastServiceCheck As DateTime = DateTime.MinValue
        Private _isServiceRunning As Boolean = False
        Private Const SERVICE_CHECK_INTERVAL_SECONDS As Integer = 120

        Friend WithEvents _toolsMenu As ToolStripMenuItem
        Friend WithEvents _checkToolsMenuItem As ToolStripMenuItem
        Friend WithEvents _updateToolsMenuItem As ToolStripMenuItem
        Friend WithEvents _monitorTasksMenuItem As ToolStripMenuItem
        Friend WithEvents _menuStrip As MenuStrip
        Friend WithEvents _fileMenu As ToolStripMenuItem
        Friend WithEvents _elevateMenuItem As ToolStripMenuItem
        Friend WithEvents _importMenuItem As ToolStripMenuItem
        Friend WithEvents _exportMenuItem As ToolStripMenuItem
        Friend WithEvents _exportAllMenuItem As ToolStripMenuItem
        Friend WithEvents _credentialManagerMenuItem As ToolStripMenuItem
        Friend WithEvents _serviceMenu As ToolStripMenuItem
        Friend WithEvents _installServiceMenuItem As ToolStripMenuItem
        Friend WithEvents _uninstallServiceMenuItem As ToolStripMenuItem
        Friend WithEvents _startServiceMenuItem As ToolStripMenuItem
        Friend WithEvents _stopServiceMenuItem As ToolStripMenuItem
        Friend WithEvents _viewMenu As ToolStripMenuItem
        Friend WithEvents _helpMenu As ToolStripMenuItem
        Friend WithEvents _optionsMenuItem As ToolStripMenuItem
        Friend WithEvents _aboutMenuItem As ToolStripMenuItem
        Private WithEvents _taskListView As ListView
        Private WithEvents _runSelectedMenuItem As ToolStripMenuItem
        Private WithEvents _editSelectedMenuItem As ToolStripMenuItem
        Private WithEvents _deleteSelectedMenuItem As ToolStripMenuItem
        Private WithEvents _exportSelectedMenuItem As ToolStripMenuItem
        Private WithEvents _checkUpdatesMenuItem As ToolStripMenuItem
        Private WithEvents _tabControl As TabControl
        Private WithEvents _createTaskButton As Button
        Private WithEvents _editTaskButton As Button
        Private WithEvents _deleteTaskButton As Button
        Private WithEvents _runTaskButton As Button
        Private WithEvents _statusStrip As StatusStrip
        Private WithEvents _statusLabel As ToolStripStatusLabel
        Private WithEvents _refreshMenuItem As ToolStripMenuItem
        'Private WithEvents _toggleLogMenuItem As ToolStripMenuItem
        Private components As System.ComponentModel.IContainer

        Public Sub New()
            Try

                ' Initialize services
                _xmlManager = ApplicationContainer.GetService(Of XMLManager)()
                _customScheduler = ApplicationContainer.GetService(Of CustomScheduler)()
                _credentialManager = ApplicationContainer.GetService(Of CredentialManager)()
                _toolManager = ApplicationContainer.GetService(Of ToolManager)()
                _logger = ApplicationContainer.GetService(Of Logger)()
                _notificationManager = ApplicationContainer.GetService(Of NotificationManager)()

                InitializeComponent()

                ' Add log event handler
                'AddHandler _logger.LogEntryAdded, AddressOf OnLogEntryAdded

                ' Initialize status update timer
                _statusTimer = New System.Windows.Forms.Timer With {
                    .Interval = 30000  ' 30 seconds
                }
                AddHandler _statusTimer.Tick, AddressOf UpdateStatus
                _statusTimer.Start()

                ' Initialize scheduler timer for GUI mode (checks tasks every 60 seconds)
                _schedulerTimer = New System.Windows.Forms.Timer With {
                    .Interval = 60000  ' 60 seconds
                }
                AddHandler _schedulerTimer.Tick, AddressOf SchedulerTimer_Tick
                _schedulerTimer.Start()

                ' Initialize auto-refresh timer
                _logger.LogInfo("MainForm initialized successfully")

                LoadSettings()
                InitializeTaskList()
                InitializeTabs()
                Me.Translate()
                RefreshTaskList()

            Catch ex As Exception
                HandleInitializationError(ex)
            End Try
        End Sub


        Private Sub AboutMenuItem_Click(sender As Object, e As EventArgs)
            MessageBox.Show($"LiteTask Version {Application.ProductVersion}", "About LiteTask", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        Private Sub AddEventsHandlers()
            'TODO: Asjust menu handlers to do the same as the buttons

            ' Buttons
            AddHandler _createTaskButton.Click, AddressOf CreateTask_Click
            AddHandler _editTaskButton.Click, AddressOf EditSelectedTasks_Click
            AddHandler _deleteTaskButton.Click, AddressOf DeleteSelectedTask_Click
            AddHandler _runTaskButton.Click, AddressOf RunSelectedTasks_Click

            ' File menu 
            AddHandler _importMenuItem.Click, AddressOf ImportTasks_Click
            AddHandler _runSelectedMenuItem.Click, AddressOf RunSelectedTasks_Click
            AddHandler _editSelectedMenuItem.Click, AddressOf EditSelectedTasks_Click
            AddHandler _deleteSelectedMenuItem.Click, AddressOf DeleteSelectedTask_Click
            AddHandler _exportSelectedMenuItem.Click, AddressOf ExportSelectedTasks_Click
            AddHandler _exportMenuItem.Click, AddressOf ExportSelectedTasks_Click
            AddHandler _exportAllMenuItem.Click, AddressOf ExportAllTasks_Click
            AddHandler _credentialManagerMenuItem.Click, AddressOf OpenCredentialManager
            AddHandler _elevateMenuItem.Click, AddressOf ElevateMenuItem_Click
            AddHandler _exitMenuItem.Click, AddressOf ExitMenuItem_Click

            ' Service menu
            AddHandler _installServiceMenuItem.Click, Sub() HandleServiceOperation("install")
            AddHandler _uninstallServiceMenuItem.Click, Sub() HandleServiceOperation("uninstall")
            AddHandler _startServiceMenuItem.Click, Sub() HandleServiceOperation("start")
            AddHandler _stopServiceMenuItem.Click, Sub() HandleServiceOperation("stop")
            'AddHandler _toggleLogMenuItem.Click, AddressOf ToggleLogVisibility

            ' View menu
            AddHandler _refreshMenuItem.Click, AddressOf RefreshMenuItem_Click

            ' Tools menu
            AddHandler _checkToolsMenuItem.Click, AddressOf CheckTools_Click
            AddHandler _updateToolsMenuItem.Click, AddressOf UpdateTools_Click
            AddHandler _monitorTasksMenuItem.Click, AddressOf MonitorTasks_Click
            AddHandler _optionsMenuItem.Click, AddressOf SettingsMenuItem_Click

            ' Help menu
            AddHandler _aboutMenuItem.Click, AddressOf AboutMenuItem_Click
            AddHandler _checkUpdatesMenuItem.Click, AddressOf CheckUpdatesMenuItem_Click

        End Sub

        'Public Sub AppendToLog(text As String)
        '    If _logTextBox.InvokeRequired Then
        '        _logTextBox.Invoke(Sub() AppendToLog(text))
        '    Else
        '        _logTextBox.AppendText(text)
        '        _logTextBox.AppendText(Environment.NewLine)
        '        _logTextBox.ScrollToCaret()
        '    End If
        'End Sub

        Private Sub AutoRefreshTimer_Tick(sender As Object, e As EventArgs)
            If _refreshPending AndAlso (DateTime.Now - _lastRefreshTime).TotalMilliseconds >= MIN_REFRESH_INTERVAL_MS Then
                _refreshPending = False
                _autoRefreshTimer.Stop()
                RefreshTaskList()
                _lastRefreshTime = DateTime.Now
            End If
        End Sub

        Private Sub CheckForUpdates()
        End Sub

        Private Sub CheckTools_Click(sender As Object, e As EventArgs)
            Dim toolStatus = _toolManager.DetectTools()
            Dim message As String = "Tool Status:" & Environment.NewLine
            For Each kvp In toolStatus
                message &= $"{kvp.Key}: {If(kvp.Value, "Present", "Missing")}" & Environment.NewLine
            Next
            MessageBox.Show(message, "Tool Check Result", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        Private Sub CheckUpdatesMenuItem_Click(sender As Object, e As EventArgs)
            ' TODO: Implement update checking logic
            MessageBox.Show("Update checking not implemented yet.", "Check for Updates", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        Private Sub CustomScheduler_TaskCompleted(sender As Object, e As TaskCompletedEventArgs)
            Try
                If IsDisposed Then Return

                If InvokeRequired Then
                    Invoke(Sub() RequestRefresh())
                Else
                    RequestRefresh()
                End If
            Catch ex As Exception
                _logger?.LogError($"Error in CustomScheduler_TaskCompleted: {ex.Message}")
            End Try
        End Sub

        Private Sub CreateTask_Click(sender As Object, e As EventArgs)

            Try
                _logger.LogInfo("Create task button clicked")

                ' For new tasks, we don't need to pass an existing task
                Using taskForm As New TaskForm(
            ApplicationContainer.GetService(Of CredentialManager)(),
            ApplicationContainer.GetService(Of CustomScheduler)(),
            ApplicationContainer.GetService(Of Logger)())

                    If taskForm.ShowDialog() = DialogResult.OK Then
                        RefreshTaskList()
                        _logger.LogInfo("Task created successfully")
                    End If
                End Using
            Catch ex As Exception
                _logger.LogError($"Error creating task: {ex.Message}")
                _logger.LogError($"StackTrace: {ex.StackTrace}")
                MessageBox.Show($"Error creating task: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        'Private Async Sub DeleteSelectedTasks(sender As Object, e As EventArgs)
        '    Try
        '        _logger.LogInfo("Delete selected tasks menu item clicked")
        '        DeleteSelectedTasksAsync()
        '    Catch ex As Exception
        '        _logger.LogError($"Error in DeleteSelectedTasks: {ex.Message}")
        '        _logger.LogError($"StackTrace: {ex.StackTrace}")
        '        MessageBox.Show($"Error deleting task(s): {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    End Try
        'End Sub

        Private Async Sub DeleteSelectedTasksAsync()
            Try
                If _taskListView.SelectedItems.Count = 0 Then
                    MessageBox.Show("Please select at least one task to delete.", "Delete Task",
                          MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim tasksToDelete = New List(Of String)()
                For Each selectedItem As ListViewItem In _taskListView.SelectedItems
                    tasksToDelete.Add(selectedItem.Text)
                Next

                Dim confirmMessage = If(tasksToDelete.Count = 1,
            $"Are you sure you want to delete the task '{tasksToDelete(0)}'?",
            $"Are you sure you want to delete these {tasksToDelete.Count} tasks?")

                If MessageBox.Show(confirmMessage, "Confirm Delete",
                         MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

                    Dim errorCount = 0
                    For Each taskName In tasksToDelete
                        Try
                            _logger.LogInfo($"Deleting task: {taskName}")
                            _customScheduler.RemoveTask(taskName)
                            _logger.LogInfo($"Successfully deleted task: {taskName}")
                        Catch ex As Exception
                            _logger.LogError($"Error deleting task {taskName}: {ex.Message}")
                            errorCount += 1
                        End Try
                    Next

                    Try
                        _customScheduler.SaveTasks()
                        _logger.LogInfo("Changes saved successfully")
                    Catch ex As Exception
                        _logger.LogError($"Error saving changes: {ex.Message}")
                        MessageBox.Show("Changes may not have been saved properly. Please verify tasks.",
                              "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End Try

                    Await RefreshUIAsync()

                    If errorCount = 0 AndAlso tasksToDelete.Count > 0 Then
                        MessageBox.Show(If(tasksToDelete.Count = 1,
                               "Task deleted successfully.",
                               $"{tasksToDelete.Count} tasks deleted successfully."),
                               "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ElseIf errorCount > 0 Then
                        MessageBox.Show($"Completed with {errorCount} error(s). Please check the log for details.",
                              "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If
                End If

            Catch ex As Exception
                _logger.LogError($"Error in DeleteSelectedTasksAsync: {ex.Message}")
                MessageBox.Show($"Error deleting tasks: {ex.Message}",
                       "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub DeleteSelectedTask_Click(sender As Object, e As EventArgs)
            Try
                _logger.LogInfo("Delete task button clicked")
                DeleteSelectedTasksAsync()
            Catch ex As Exception
                _logger.LogError($"Error in DeleteTaskButton_Click: {ex.Message}")
                MessageBox.Show($"Error deleting task(s): {ex.Message}",
                       "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        'Private Async Function EditTaskAsync(_task As ScheduledTask) As Task
        '    Try
        '        _logger.LogInfo($"Starting task edit for task: {_task.Name}")

        '        ' Create a deep copy of the task for editing
        '        Dim taskCopy = _task.Clone()

        '        Using taskForm As New TaskForm(_credentialManager, _customScheduler, _logger, taskCopy)
        '            If taskForm.ShowDialog() = DialogResult.OK Then
        '                Try
        '                    Await Task.Run(Sub()
        '                                       _customScheduler.UpdateTask(taskCopy)
        '                                       _customScheduler.SaveTasks()
        '                                   End Sub)
        '                    Await RefreshUIAsync()
        '                    _logger.LogInfo($"Task {_task.Name} edited and saved successfully")
        '                    MessageBox.Show("Task updated successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        '                Catch ex As Exception
        '                    _logger.LogError($"Error saving edited task {_task.Name}: {ex.Message}")
        '                    Throw New Exception($"Failed to save task changes: {ex.Message}", ex)
        '                End Try
        '            End If
        '        End Using
        '    Catch ex As Exception
        '        _logger.LogError($"Error editing task: {ex.Message}")
        '        _logger.LogError($"StackTrace: {ex.StackTrace}")
        '        MessageBox.Show($"Error editing task: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    End Try
        'End Function

        Private Sub EditSelectedTasks_Click(sender As Object, e As EventArgs)
            Try
                _logger.LogInfo("Edit task button clicked")

                If _taskListView.SelectedItems.Count = 0 Then
                    MessageBox.Show("Please select a task to edit.", "Edit Task", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim selectedItem = _taskListView.SelectedItems(0)
                Dim taskName = selectedItem.Text
                Dim task = _customScheduler.GetTask(taskName)  ' Get fresh task from scheduler

                If task Is Nothing Then
                    _logger.LogError($"Task not found: {taskName}")
                    MessageBox.Show("Selected task could not be loaded.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                Using taskForm As New TaskForm(_credentialManager, _customScheduler, _logger, task)
                    If taskForm.ShowDialog() = DialogResult.OK Then
                        RefreshTaskList()
                        _logger.LogInfo($"Task {taskName} edited successfully")
                    End If
                End Using

            Catch ex As Exception
                _logger.LogError($"Error in EditTaskButton_Click: {ex.Message}")
                MessageBox.Show($"Error editing task: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        'Private Async Sub EditSelectedTask(sender As Object, e As EventArgs)
        '    Try
        '        If _taskListView.SelectedItems.Count <> 1 Then
        '            MessageBox.Show("Please select a single task to edit.", "Edit Task", MessageBoxButtons.OK, MessageBoxIcon.Information)
        '            Return
        '        End If

        '        Dim selectedItem = _taskListView.SelectedItems(0)
        '        Dim taskName = selectedItem.Text
        '        Dim task = _customScheduler.GetTask(taskName)

        '        If task IsNot Nothing Then
        '            Using taskForm As New TaskForm(_credentialManager, _customScheduler, _logger, task)
        '                If taskForm.ShowDialog() = DialogResult.OK Then
        '                    Await RefreshUIAsync()
        '                End If
        '            End Using
        '        End If
        '    Catch ex As Exception
        '        _logger.LogError($"Error editing task: {ex.Message}")
        '        MessageBox.Show($"Error editing task: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    End Try
        'End Sub

        Private Sub ElevateMenuItem_Click(sender As Object, e As EventArgs)
            If Not IsElevated() Then
                If MessageBox.Show(
            TranslationManager.Instance.GetTranslation("ElevateConfirm.Message"),
            TranslationManager.Instance.GetTranslation("ElevateConfirm.Title"),
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question) = DialogResult.Yes Then

                    ' Save current state
                    _customScheduler.SaveTasks()

                    ' Restart elevated
                    RestartAsAdmin("")
                End If
            End If
        End Sub


        Private Sub ExecuteServiceCommand(command As String, operationType As String, logPath As String)
            Try
                ' Create process to execute command
                Dim process As New Process()
                process.StartInfo.FileName = "cmd.exe"
                process.StartInfo.Arguments = $"/c {command}"
                process.StartInfo.UseShellExecute = False
                process.StartInfo.RedirectStandardOutput = True
                process.StartInfo.RedirectStandardError = True
                process.StartInfo.CreateNoWindow = True

                ' Log the command execution (with timestamp)
                Dim logEntry As String = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {operationType}{Environment.NewLine}"
                logEntry += $"Command: {command}{Environment.NewLine}"
                File.AppendAllText(logPath, logEntry)

                ' Execute the command
                process.Start()
                Dim output As String = process.StandardOutput.ReadToEnd()
                Dim errors As String = process.StandardError.ReadToEnd()
                process.WaitForExit()

                ' Log the results
                logEntry = $"Output: {output}{Environment.NewLine}"
                If Not String.IsNullOrEmpty(errors) Then
                    logEntry += $"Errors: {errors}{Environment.NewLine}"
                End If
                logEntry += $"Exit Code: {process.ExitCode}{Environment.NewLine}"
                logEntry += New String("-", 50) & Environment.NewLine

                File.AppendAllText(logPath, logEntry)

                If process.ExitCode <> 0 Then
                    Throw New Exception($"Command failed with exit code {process.ExitCode}. Error: {errors}")
                End If

                _logger.LogInfo($"Service command executed successfully: {command}")
                _logger.LogInfo($"Command output: {output}")
            Catch ex As Exception
                _logger.LogError($"Error executing service command: {ex.Message}")
                Throw
            End Try

        End Sub

        Private Sub ExitMenuItem_Click(sender As Object, e As EventArgs)
            If MessageBox.Show(
            TranslationManager.Instance.GetTranslation("ExitConfirm.Message"),
            TranslationManager.Instance.GetTranslation("ExitConfirm.Title"),
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question) = DialogResult.Yes Then
                Application.Exit()
            End If
        End Sub

        Private Sub ExportAllTasks_Click(sender As Object, e As EventArgs)
            Try
                Using saveFileDialog As New SaveFileDialog()
                    saveFileDialog.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
                    saveFileDialog.FilterIndex = 1
                    saveFileDialog.RestoreDirectory = True

                    If saveFileDialog.ShowDialog() = DialogResult.OK Then
                        ExportTasks(saveFileDialog.FileName, GetAllTasks())
                    End If
                End Using
            Catch ex As Exception
                'UpdateLog($"Error: {ex.Message}")
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Public Sub ExportTasks(filePath As String, taskNames As List(Of String))
            Try
                _logger.LogInfo($"Exporting tasks to file: {filePath}")
                If String.IsNullOrEmpty(filePath) Then
                    Throw New ArgumentException("File path cannot be null or empty.", NameOf(filePath))
                End If
                If taskNames Is Nothing OrElse taskNames.Count = 0 Then
                    Throw New ArgumentException("Task names list cannot be null or empty.", NameOf(taskNames))
                End If

                Dim exportXmlManager As New XMLManager(filePath)

                For Each taskName In taskNames
                    If String.IsNullOrEmpty(taskName) Then
                        _logger.LogWarning("Encountered a null or empty task name during export")
                        Continue For
                    End If

                    Dim task = _customScheduler.GetTask(taskName)
                    If task IsNot Nothing Then
                        exportXmlManager.SaveTask(task)
                        _logger.LogInfo($"Exported task: {taskName}")
                    Else
                        _logger.LogWarning($"Task not found for export: {taskName}")
                    End If
                Next

                _logger.LogInfo($"Exported {taskNames.Count} tasks successfully")
            Catch ex As Exception
                _logger.LogError($"Error exporting tasks: {ex.Message}")
                _logger.LogError($"StackTrace: {ex.StackTrace}")
                Throw
            End Try
        End Sub

        Private Sub ExportSelectedTasks_Click(sender As Object, e As EventArgs)
            Try

                Dim selectedTasks = GetSelectedTasks()
                If selectedTasks Is Nothing OrElse selectedTasks.Count = 0 Then
                    MessageBox.Show("Please select at least one task to export.", "No Tasks Selected", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Using saveFileDialog As New SaveFileDialog()
                    saveFileDialog.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
                    saveFileDialog.FilterIndex = 1
                    saveFileDialog.RestoreDirectory = True

                    If saveFileDialog.ShowDialog() = DialogResult.OK Then
                        ExportTasks(saveFileDialog.FileName, selectedTasks)
                        MessageBox.Show($"Successfully exported {selectedTasks.Count} task(s).", "Export Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                End Using
            Catch ex As Exception
                _logger.LogError($"Error exporting tasks: {ex.Message}")
                _logger.LogError($"StackTrace: {ex.StackTrace}")
                MessageBox.Show($"Error exporting tasks: {ex.Message}", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Public Function GetAllTasks() As List(Of String)
            Return _customScheduler.GetAllTasks().Select(Function(t) t.Name).ToList()
        End Function

        Public Function GetSelectedTasks() As List(Of String)
            If _taskListView Is Nothing Then
                _logger.LogError("TaskListView is not initialized in GetSelectedTasks")
                Return New List(Of String)()
            End If

            Return _taskListView.SelectedItems.Cast(Of ListViewItem)().
                Select(Function(item) item.Text).
                Where(Function(taskName) Not String.IsNullOrEmpty(taskName)).
                ToList()
        End Function

        Private Function GetTaskStatus(task As ScheduledTask) As String
            If task.Actions.Any(Function(a) a.Status = TaskActionStatus.Running) Then
                Return "Running"
            ElseIf task.Actions.Any(Function(a) a.Status = TaskActionStatus.Failed) Then
                Return "Failed"
            ElseIf task.Actions.All(Function(a) a.Status = TaskActionStatus.Completed) Then
                Return "Completed"
            Else
                Return "Pending"
            End If
        End Function

        Private Sub HandleInitializationError(ex As Exception)
            Dim errorMessage = $"Error initializing MainForm: {ex.Message}"

            If _logger IsNot Nothing Then
                _logger.LogError(errorMessage, ex)
            Else
                Debug.WriteLine(errorMessage)
                Debug.WriteLine(ex.StackTrace)
            End If

            MessageBox.Show(errorMessage, "Initialization Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Sub

        Private Sub HandleServiceOperation(operation As String)
            If Not IsElevated() Then
                RestartAsAdmin($"-{operation}")
                Return
            End If

            Dim serviceName As String = "LiteTaskService"
            Dim exePath As String = Application.ExecutablePath
            Dim logPath As String = Path.Combine(Application.StartupPath, "LiteTaskData", "logs", "service.log")

            ' Ensure log directory exists
            Directory.CreateDirectory(Path.GetDirectoryName(logPath))

            Try
                Select Case operation.ToLower()
                    Case "install"
                        ExecuteServiceCommand($"sc create {serviceName} binPath= ""{exePath} -service"" start= auto", "Service Installation", logPath)
                        MessageBox.Show($"Service {serviceName} installed.", "Service Installation", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    Case "uninstall"
                        ExecuteServiceCommand($"sc delete {serviceName}", "Service Uninstallation", logPath)
                        MessageBox.Show($"Service {serviceName} uninstalled.", "Service Uninstallation", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    Case "start"
                        ExecuteServiceCommand($"sc start {serviceName}", "Service Start", logPath)
                        MessageBox.Show($"Service {serviceName} started.", "Service Start", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    Case "stop"
                        ExecuteServiceCommand($"sc stop {serviceName}", "Service Stop", logPath)
                        MessageBox.Show($"Service {serviceName} stopped.", "Service Stop", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    Case Else
                        MessageBox.Show("Invalid operation.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Select
            Catch ex As Exception
                _logger.LogError($"Error in HandleServiceOperation: {ex.Message}")
                MessageBox.Show($"Error performing service operation: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Public Sub ImportTasks(filePath As String)
            Try
                _logger.LogInfo($"Importing tasks from file: {filePath}")
                Dim importXmlManager As New XMLManager(filePath)
                Dim importedTaskNames = importXmlManager.GetAllTaskNames()

                For Each taskName In importedTaskNames
                    Dim task = importXmlManager.LoadTask(taskName)

                    ' Check if a task with the same name already exists
                    If _customScheduler.GetTask(taskName) IsNot Nothing Then
                        Dim result = MessageBox.Show($"A task with the name '{taskName}' already exists. Do you want to overwrite it?",
                                                 "Task Already Exists",
                                                 MessageBoxButtons.YesNoCancel,
                                                 MessageBoxIcon.Question)

                        Select Case result
                            Case DialogResult.Yes
                                _customScheduler.UpdateTask(task)
                            Case DialogResult.No
                                Continue For
                            Case DialogResult.Cancel
                                _logger.LogInfo("Task import cancelled by user")
                                Return
                        End Select
                    Else
                        If String.IsNullOrEmpty(task.AccountType) Then
                            task.AccountType = _defaultAccountType
                        End If
                        _customScheduler.AddTask(task)
                    End If

                    _logger.LogInfo($"Imported task: {taskName}")
                Next

                MessageBox.Show($"Successfully imported {importedTaskNames.Count} tasks.", "Import Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
                _logger.LogInfo($"Imported {importedTaskNames.Count} tasks successfully")
            Catch ex As Exception
                _logger.LogError($"Error importing tasks: {ex.Message}")
                _logger.LogError($"StackTrace: {ex.StackTrace}")
                MessageBox.Show($"Error importing tasks: {ex.Message}", "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub ImportTasks_Click(sender As Object, e As EventArgs)
            Try
                Using openFileDialog As New OpenFileDialog()
                    openFileDialog.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
                    openFileDialog.FilterIndex = 1
                    openFileDialog.RestoreDirectory = True

                    If openFileDialog.ShowDialog() = DialogResult.OK Then
                        ImportTasks(openFileDialog.FileName)
                    End If
                End Using
            Catch ex As Exception
                'UpdateLog($"Error: {ex.Message}")
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub InitializeComponent()
            components = New Container()
            Dim resources As ComponentResourceManager = New ComponentResourceManager(GetType(MainForm))
            _tabControl = New TabControl()
            _taskTabPage = New TabPage()
            _taskTableLayout = New TableLayoutPanel()
            _taskListView = New ListView()
            _contextMenuStrip = New ContextMenuStrip(components)
            _runSelectedMenuItem = New ToolStripMenuItem()
            _editSelectedMenuItem = New ToolStripMenuItem()
            _deleteSelectedMenuItem = New ToolStripMenuItem()
            _exportSelectedMenuItem = New ToolStripMenuItem()
            _buttonPanel = New FlowLayoutPanel()
            _runTaskButton = New Button()
            _createTaskButton = New Button()
            _editTaskButton = New Button()
            _deleteTaskButton = New Button()
            _menuStrip = New MenuStrip()
            _fileMenu = New ToolStripMenuItem()
            _importMenuItem = New ToolStripMenuItem()
            _exportMenuItem = New ToolStripMenuItem()
            _exportAllMenuItem = New ToolStripMenuItem()
            _credentialManagerMenuItem = New ToolStripMenuItem()
            _elevateMenuItem = New ToolStripMenuItem()
            _exitMenuItem = New ToolStripMenuItem()
            _serviceMenu = New ToolStripMenuItem()
            _installServiceMenuItem = New ToolStripMenuItem()
            _uninstallServiceMenuItem = New ToolStripMenuItem()
            _startServiceMenuItem = New ToolStripMenuItem()
            _stopServiceMenuItem = New ToolStripMenuItem()
            _viewMenu = New ToolStripMenuItem()
            _refreshMenuItem = New ToolStripMenuItem()
            '_toggleLogMenuItem = New ToolStripMenuItem()
            _toolsMenu = New ToolStripMenuItem()
            _checkToolsMenuItem = New ToolStripMenuItem()
            _updateToolsMenuItem = New ToolStripMenuItem()
            _monitorTasksMenuItem = New ToolStripMenuItem()
            _optionsMenuItem = New ToolStripMenuItem()
            _helpMenu = New ToolStripMenuItem()
            _aboutMenuItem = New ToolStripMenuItem()
            _checkUpdatesMenuItem = New ToolStripMenuItem()
            _splitContainer = New SplitContainer()
            '_logTextBox = New TextBox()
            _taskPanel = New Panel()
            _statusStrip = New StatusStrip()
            _statusLabel = New ToolStripStatusLabel()
            _tabControl.SuspendLayout()
            _taskTabPage.SuspendLayout()
            _taskTableLayout.SuspendLayout()
            _contextMenuStrip.SuspendLayout()
            _buttonPanel.SuspendLayout()
            _menuStrip.SuspendLayout()
            CType(_splitContainer, ISupportInitialize).BeginInit()
            _splitContainer.Panel1.SuspendLayout()
            '_splitContainer.Panel2.SuspendLayout()
            _splitContainer.SuspendLayout()
            _statusStrip.SuspendLayout()
            SuspendLayout()
            ' 
            ' _tabControl
            ' 
            _tabControl.Controls.Add(_taskTabPage)
            _tabControl.Dock = DockStyle.Fill
            _tabControl.Location = New Point(0, 0)
            _tabControl.Margin = New Padding(3, 2, 3, 2)
            _tabControl.Name = "_tabControl"
            _tabControl.SelectedIndex = 0
            _tabControl.Size = New Size(701, 376)
            _tabControl.TabIndex = 0
            ' 
            ' _taskTabPage
            ' 
            _taskTabPage.Controls.Add(_taskTableLayout)
            _taskTabPage.Location = New Point(4, 24)
            _taskTabPage.Margin = New Padding(3, 2, 3, 2)
            _taskTabPage.Name = "_taskTabPage"
            _taskTabPage.Size = New Size(693, 348)
            _taskTabPage.TabIndex = 0
            _taskTabPage.Text = "Tasks"
            ' 
            ' _taskTableLayout
            ' 
            _taskTableLayout.ColumnCount = 1
            _taskTableLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 694.0F))
            _taskTableLayout.Controls.Add(_taskListView, 0, 0)
            _taskTableLayout.Controls.Add(_buttonPanel, 0, 1)
            _taskTableLayout.Dock = DockStyle.Fill
            _taskTableLayout.Location = New Point(0, 0)
            _taskTableLayout.Margin = New Padding(3, 2, 3, 2)
            _taskTableLayout.Name = "_taskTableLayout"
            _taskTableLayout.RowCount = 2
            _taskTableLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 90.0F))
            _taskTableLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 10.0F))
            _taskTableLayout.Size = New Size(693, 348)
            _taskTableLayout.TabIndex = 0
            ' 
            ' _taskListView
            ' 
            _taskListView.ContextMenuStrip = _contextMenuStrip
            _taskListView.Dock = DockStyle.Fill
            _taskListView.FullRowSelect = True
            _taskListView.Location = New Point(3, 2)
            _taskListView.Margin = New Padding(3, 2, 3, 2)
            _taskListView.Name = "_taskListView"
            _taskListView.Size = New Size(688, 309)
            _taskListView.TabIndex = 0
            _taskListView.UseCompatibleStateImageBehavior = False
            _taskListView.View = View.Details
            ' 
            ' _contextMenuStrip
            ' 
            _contextMenuStrip.ImageScalingSize = New Size(20, 20)
            _contextMenuStrip.Items.AddRange(New ToolStripItem() {_runSelectedMenuItem, _editSelectedMenuItem, _deleteSelectedMenuItem, _exportSelectedMenuItem})
            _contextMenuStrip.Name = "_contextMenuStrip"
            _contextMenuStrip.Size = New Size(194, 92)
            ' 
            ' _runSelectedMenuItem
            ' 
            _runSelectedMenuItem.Name = "_runSelectedMenuItem"
            _runSelectedMenuItem.Size = New Size(193, 22)
            _runSelectedMenuItem.Text = "Run Selected Task(s)"
            ' 
            ' _editSelectedMenuItem
            ' 
            _editSelectedMenuItem.Name = "_editSelectedMenuItem"
            _editSelectedMenuItem.Size = New Size(193, 22)
            _editSelectedMenuItem.Text = "Edit Selected Task"
            ' 
            ' _deleteSelectedMenuItem
            ' 
            _deleteSelectedMenuItem.Name = "_deleteSelectedMenuItem"
            _deleteSelectedMenuItem.Size = New Size(193, 22)
            _deleteSelectedMenuItem.Text = "Delete Selected Task(s)"
            ' 
            ' _exportSelectedMenuItem
            ' 
            _exportSelectedMenuItem.Name = "_exportSelectedMenuItem"
            _exportSelectedMenuItem.Size = New Size(193, 22)
            _exportSelectedMenuItem.Text = "Export Selected Task(s)"
            ' 
            ' _buttonPanel
            ' 
            _buttonPanel.AutoSize = True
            _buttonPanel.Controls.Add(_runTaskButton)
            _buttonPanel.Controls.Add(_createTaskButton)
            _buttonPanel.Controls.Add(_editTaskButton)
            _buttonPanel.Controls.Add(_deleteTaskButton)
            _buttonPanel.Dock = DockStyle.Fill
            _buttonPanel.Location = New Point(3, 315)
            _buttonPanel.Margin = New Padding(3, 2, 3, 2)
            _buttonPanel.Name = "_buttonPanel"
            _buttonPanel.Padding = New Padding(9, 4, 9, 4)
            _buttonPanel.Size = New Size(688, 31)
            _buttonPanel.TabIndex = 1
            _buttonPanel.WrapContents = False
            ' 
            ' _runTaskButton
            ' 
            _runTaskButton.Location = New Point(12, 6)
            _runTaskButton.Margin = New Padding(3, 2, 3, 2)
            _runTaskButton.Name = "_runTaskButton"
            _runTaskButton.Size = New Size(88, 25)
            _runTaskButton.TabIndex = 0
            _runTaskButton.Text = "Run task"
            ' 
            ' _createTaskButton
            ' 
            _createTaskButton.Location = New Point(106, 6)
            _createTaskButton.Margin = New Padding(3, 2, 3, 2)
            _createTaskButton.Name = "_createTaskButton"
            _createTaskButton.Size = New Size(88, 25)
            _createTaskButton.TabIndex = 1
            _createTaskButton.Text = "Create task"
            ' 
            ' _editTaskButton
            ' 
            _editTaskButton.Location = New Point(200, 6)
            _editTaskButton.Margin = New Padding(3, 2, 3, 2)
            _editTaskButton.Name = "_editTaskButton"
            _editTaskButton.Size = New Size(88, 25)
            _editTaskButton.TabIndex = 2
            _editTaskButton.Text = "Edit task"
            ' 
            ' _deleteTaskButton
            ' 
            _deleteTaskButton.Location = New Point(294, 6)
            _deleteTaskButton.Margin = New Padding(3, 2, 3, 2)
            _deleteTaskButton.Name = "_deleteTaskButton"
            _deleteTaskButton.Size = New Size(88, 25)
            _deleteTaskButton.TabIndex = 3
            _deleteTaskButton.Text = "Delete task"
            ' 
            ' _menuStrip
            ' 
            _menuStrip.ImageScalingSize = New Size(20, 20)
            _menuStrip.Items.AddRange(New ToolStripItem() {_fileMenu, _serviceMenu, _viewMenu, _toolsMenu, _helpMenu})
            _menuStrip.Location = New Point(0, 0)
            _menuStrip.Name = "_menuStrip"
            _menuStrip.Padding = New Padding(5, 2, 0, 2)
            _menuStrip.Size = New Size(701, 24)
            _menuStrip.TabIndex = 1
            _menuStrip.Text = "menuStrip1"
            ' 
            ' _fileMenu
            ' 
            _fileMenu.DropDownItems.AddRange(New ToolStripItem() {
                _importMenuItem,
                _exportMenuItem,
                _exportAllMenuItem,
                New ToolStripSeparator(),  ' Spacer
                _credentialManagerMenuItem,
                New ToolStripSeparator(),  ' Spacer
                _elevateMenuItem,
                New ToolStripSeparator(),  ' Spacer
                _exitMenuItem})
            '
            _fileMenu.Name = "_fileMenu"
            _fileMenu.Size = New Size(37, 20)
            _fileMenu.Text = "File"
            ' 
            ' _importMenuItem
            ' 
            _importMenuItem.Name = "_importMenuItem"
            _importMenuItem.Size = New Size(180, 22)
            _importMenuItem.Text = "Import Tasks"
            ' 
            ' _exportMenuItem
            ' 
            _exportMenuItem.Name = "_exportMenuItem"
            _exportMenuItem.Size = New Size(180, 22)
            _exportMenuItem.Text = "Export Tasks"
            ' 
            ' _exportAllMenuItem
            ' 
            _exportAllMenuItem.Name = "_exportAllMenuItem"
            _exportAllMenuItem.Size = New Size(180, 22)
            _exportAllMenuItem.Text = "Export All Tasks"


            ' 
            ' _credentialManagerMenuItem
            ' 
            _credentialManagerMenuItem.Name = "_credentialManagerMenuItem"
            _credentialManagerMenuItem.Size = New Size(180, 22)
            _credentialManagerMenuItem.Text = "Credential Manager"


            ' _elevateMenuItem
            _elevateMenuItem.Name = "_elevateMenuItem"
            _elevateMenuItem.Size = New Size(180, 22)
            _elevateMenuItem.Text = "Run as Administrator"

            ' _exitMenuItem
            _exitMenuItem.Name = "_exitMenuItem"
            _exitMenuItem.Size = New Size(180, 22)
            _exitMenuItem.Text = "Exit LiteTask"
            ' 
            ' _serviceMenu
            ' 
            _serviceMenu.DropDownItems.AddRange(New ToolStripItem() {_installServiceMenuItem, _uninstallServiceMenuItem, _startServiceMenuItem, _stopServiceMenuItem})
            _serviceMenu.Name = "_serviceMenu"
            _serviceMenu.Size = New Size(56, 20)
            _serviceMenu.Text = "Service"
            ' 
            ' _installServiceMenuItem
            ' 
            _installServiceMenuItem.Name = "_installServiceMenuItem"
            _installServiceMenuItem.Size = New Size(180, 22)
            _installServiceMenuItem.Text = "Install Service"
            ' 
            ' _uninstallServiceMenuItem
            ' 
            _uninstallServiceMenuItem.Name = "_uninstallServiceMenuItem"
            _uninstallServiceMenuItem.Size = New Size(180, 22)
            _uninstallServiceMenuItem.Text = "Uninstall Service"
            ' 
            ' _startServiceMenuItem
            ' 
            _startServiceMenuItem.Name = "_startServiceMenuItem"
            _startServiceMenuItem.Size = New Size(180, 22)
            _startServiceMenuItem.Text = "Start Service"
            ' 
            ' _stopServiceMenuItem
            ' 
            _stopServiceMenuItem.Name = "_stopServiceMenuItem"
            _stopServiceMenuItem.Size = New Size(180, 22)
            _stopServiceMenuItem.Text = "Stop Service"
            ' 
            ' _viewMenu
            ' 
            _viewMenu.DropDownItems.AddRange(New ToolStripItem() {_refreshMenuItem})
            _viewMenu.Name = "_viewMenu"
            _viewMenu.Size = New Size(44, 20)
            _viewMenu.Text = "View"
            ' 
            ' _refreshMenuItem
            ' 
            _refreshMenuItem.Name = "_refreshMenuItem"
            _refreshMenuItem.Size = New Size(180, 22)
            _refreshMenuItem.Text = "Refresh"
            ' 
            '' _toggleLogMenuItem
            '' 
            '_toggleLogMenuItem.Name = "_toggleLogMenuItem"
            '_toggleLogMenuItem.Size = New Size(180, 22)
            '_toggleLogMenuItem.Text = "Show/Hide Log"
            '' 
            ' _toolsMenu
            ' 
            _toolsMenu.DropDownItems.AddRange(New ToolStripItem() {_checkToolsMenuItem,
                _updateToolsMenuItem,
                New ToolStripSeparator(),  ' Spacer
                _monitorTasksMenuItem,
                New ToolStripSeparator(),  ' Spacer
                _optionsMenuItem})

            _toolsMenu.Name = "_toolsMenu"
            _toolsMenu.Size = New Size(46, 20)
            _toolsMenu.Text = "Tools"
            ' 
            ' _checkToolsMenuItem
            ' 
            _checkToolsMenuItem.Name = "_checkToolsMenuItem"
            _checkToolsMenuItem.Size = New Size(215, 22)
            _checkToolsMenuItem.Text = "Check Tools"
            ' 
            ' _updateToolsMenuItem
            ' 
            _updateToolsMenuItem.Name = "_updateToolsMenuItem"
            _updateToolsMenuItem.Size = New Size(215, 22)
            _updateToolsMenuItem.Text = "Update Tools"
            ' 
            '
            ' _monitorTasksMenuItem
            ' 
            _monitorTasksMenuItem.Name = "_monitorTasksMenuItem"
            _monitorTasksMenuItem.Size = New Size(215, 22)
            _monitorTasksMenuItem.Text = "Monitor LiteTask Processes"
            '
            ' 
            ' _optionsMenuItem
            ' 
            _optionsMenuItem.Name = "_optionsMenuItem"
            _optionsMenuItem.Size = New Size(215, 22)
            _optionsMenuItem.Text = "Options"
            ' 
            ' _helpMenu
            ' 
            _helpMenu.DropDownItems.AddRange(New ToolStripItem() {_aboutMenuItem, _checkUpdatesMenuItem})
            _helpMenu.Name = "_helpMenu"
            _helpMenu.Size = New Size(44, 20)
            _helpMenu.Text = "Help"
            ' 
            ' _aboutMenuItem
            ' 
            _aboutMenuItem.Name = "_aboutMenuItem"
            _aboutMenuItem.Size = New Size(180, 22)
            _aboutMenuItem.Text = "About"
            ' 
            ' _checkUpdatesMenuItem
            ' 
            _checkUpdatesMenuItem.Name = "_checkUpdatesMenuItem"
            _checkUpdatesMenuItem.Size = New Size(180, 22)
            _checkUpdatesMenuItem.Text = "Check for Updates"
            ' 
            ' _splitContainer
            ' 
            _splitContainer.Dock = DockStyle.Fill
            _splitContainer.Location = New Point(0, 24)
            _splitContainer.Margin = New Padding(3, 2, 3, 2)
            _splitContainer.Name = "_splitContainer"
            _splitContainer.Orientation = Orientation.Horizontal
            ' 
            ' _splitContainer.Panel1
            ' 
            _splitContainer.Panel1.Controls.Add(_tabControl)
            ' 
            '' _splitContainer.Panel2
            '' 
            '_splitContainer.Panel2.Controls.Add(_logTextBox)
            '_splitContainer.Panel2Collapsed = True
            '_splitContainer.Size = New Size(701, 376)
            '_splitContainer.SplitterDistance = 38
            '_splitContainer.SplitterWidth = 3
            '_splitContainer.TabIndex = 1
            ' 
            ' _logTextBox
            ' 
            '_logTextBox.Dock = DockStyle.Fill
            '_logTextBox.Location = New Point(0, 0)
            '_logTextBox.Margin = New Padding(3, 2, 3, 2)
            '_logTextBox.Multiline = True
            '_logTextBox.Name = "_logTextBox"
            '_logTextBox.ReadOnly = True
            '_logTextBox.ScrollBars = ScrollBars.Vertical
            '_logTextBox.Size = New Size(131, 34)
            '_logTextBox.TabIndex = 0
            ' 
            ' _taskPanel
            ' 
            _taskPanel.Location = New Point(0, 0)
            _taskPanel.Name = "_taskPanel"
            _taskPanel.Size = New Size(200, 100)
            _taskPanel.TabIndex = 0
            ' 
            ' _statusStrip
            ' 
            _statusStrip.ImageScalingSize = New Size(20, 20)
            _statusStrip.Items.AddRange(New ToolStripItem() {_statusLabel})
            _statusStrip.Location = New Point(0, 400)
            _statusStrip.Name = "_statusStrip"
            _statusStrip.Padding = New Padding(1, 0, 12, 0)
            _statusStrip.Size = New Size(701, 22)
            _statusStrip.TabIndex = 2
            _statusStrip.Text = "statusStrip1"
            ' 
            ' _statusLabel
            ' 
            _statusLabel.Name = "_statusLabel"
            _statusLabel.Size = New Size(39, 17)
            _statusLabel.Text = "Ready"
            ' 
            ' MainForm
            ' 
            AutoScaleDimensions = New SizeF(7.0F, 15.0F)
            AutoScaleMode = AutoScaleMode.Font
            ClientSize = New Size(701, 422)
            Controls.Add(_splitContainer)
            Controls.Add(_menuStrip)
            Controls.Add(_statusStrip)
            Icon = CType(resources.GetObject("$this.Icon"), Icon)
            Margin = New Padding(3, 2, 3, 2)
            MinimumSize = New Size(702, 460)
            Name = "MainForm"
            Text = "LiteTask"
            _tabControl.ResumeLayout(False)
            _taskTabPage.ResumeLayout(False)
            _taskTableLayout.ResumeLayout(False)
            _taskTableLayout.PerformLayout()
            _contextMenuStrip.ResumeLayout(False)
            _buttonPanel.ResumeLayout(False)
            _menuStrip.ResumeLayout(False)
            _menuStrip.PerformLayout()
            _splitContainer.Panel1.ResumeLayout(False)
            '_splitContainer.Panel2.ResumeLayout(False)
            '_splitContainer.Panel2.PerformLayout()
            CType(_splitContainer, ISupportInitialize).EndInit()
            _splitContainer.Panel2Collapsed = True
            _splitContainer.ResumeLayout(False)
            _statusStrip.ResumeLayout(False)
            _statusStrip.PerformLayout()
            ResumeLayout(False)
            PerformLayout()

            ' Update Tabpage text
            _taskTabPage.Text = TranslationManager.Instance.GetTranslation("_taskTabPage")

            ' Update button texts
            _runTaskButton.Text = TranslationManager.Instance.GetTranslation("_runTaskButton.Text")
            _createTaskButton.Text = TranslationManager.Instance.GetTranslation("_createTaskButton.Text")
            _editTaskButton.Text = TranslationManager.Instance.GetTranslation("_editTaskButton.Text")
            _deleteTaskButton.Text = TranslationManager.Instance.GetTranslation("_deleteTaskButton.Text")

            AddEventsHandlers()
            InitializeAutoRefresh()
        End Sub

        Private Sub InitializeAutoRefresh()
            _autoRefreshTimer = New System.Windows.Forms.Timer With {
        .Interval = MIN_REFRESH_INTERVAL_MS,
        .Enabled = False
            }
            AddHandler _autoRefreshTimer.Tick, AddressOf AutoRefreshTimer_Tick
            AddHandler _customScheduler.TaskCompleted, AddressOf CustomScheduler_TaskCompleted
        End Sub

        Private Sub InitializeTabs()
            Try
                _logger.LogInfo("Initializing tabs")

                If _tabControl Is Nothing Then
                    _logger.LogError("_tabControl is null")
                    Return
                End If

                ' Add Tabs
                _logger.LogInfo("Clearing tab pages")
                _tabControl.TabPages.Clear()

                If _taskTabPage Is Nothing Then
                    _logger.LogError("_taskTabPage is null")
                    Return
                End If

                _logger.LogInfo("Adding task tab page")
                _tabControl.TabPages.Add(_taskTabPage)

                ' Add PowerShell tab
                _logger.LogInfo("Attempting to get RunTab service")
                Dim runTab = ApplicationContainer.GetService(Of RunTab)()
                If runTab Is Nothing Then
                    _logger.LogWarning("RunTab service returned null")
                Else
                    _logger.LogInfo("Getting RunTab page")
                    Dim runTabPage = runTab.GetTabPage()
                    If runTabPage Is Nothing Then
                        _logger.LogWarning("RunTab.GetTabPage() returned null")
                    Else
                        _logger.LogInfo("Adding RunTab page to tabs")
                        _tabControl.TabPages.Add(runTabPage)
                        Try
                            _logger.LogInfo("Attempting to translate RunTab")
                            runTab.TranslateRunTab()
                            _logger.LogInfo("RunTab translated successfully")
                        Catch translationEx As Exception
                            _logger.LogWarning($"Failed to translate RunTab: {translationEx.Message}")
                        End Try
                    End If
                End If

                ' Add SQL tab
                _logger.LogInfo("Attempting to get SqlTab service")
                Dim sqlTab = ApplicationContainer.GetService(Of SqlTab)()
                If sqlTab Is Nothing Then
                    _logger.LogWarning("SqlTab service returned null")
                Else
                    _logger.LogInfo("Getting SqlTab page")
                    Dim sqlTabPage = sqlTab.GetTabPage()
                    If sqlTabPage Is Nothing Then
                        _logger.LogWarning("SqlTab.GetTabPage() returned null")
                    Else
                        _logger.LogInfo("Adding SqlTab page to tabs")
                        _tabControl.TabPages.Add(sqlTabPage)
                        Try
                            _logger.LogInfo("Attempting to translate SqlTab")
                            sqlTab.TranslateSqlTab()
                            _logger.LogInfo("SqlTab translated successfully")
                        Catch translationEx As Exception
                            _logger.LogWarning($"Failed to translate SqlTab: {translationEx.Message}")
                        End Try
                    End If
                End If

                _logger.LogInfo("Tabs initialization completed")
            Catch ex As Exception
                _logger.LogError($"Error in InitializeTabs: {ex.Message}")
                _logger.LogError($"Stack Trace: {ex.StackTrace}")
                If ex.InnerException IsNot Nothing Then
                    _logger.LogError($"Inner Exception: {ex.InnerException.Message}")
                    _logger.LogError($"Inner Stack Trace: {ex.InnerException.StackTrace}")
                End If
                MessageBox.Show($"An error occurred while initializing tabs: {ex.Message}",
               "Initialization Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub InitializeTaskList()
            Try
                _logger.LogInfo("Initializing task list")

                ' Configure ListView columns for tasks
                _taskListView.Clear()
                _taskListView.Columns.Clear()
                _taskListView.View = View.Details
                _taskListView.FullRowSelect = True
                _taskListView.GridLines = True
                _taskListView.MultiSelect = True
                _taskListView.HideSelection = False

                ' Add columns with appropriate widths
                _taskListView.Columns.AddRange(New ColumnHeader() {
                New ColumnHeader With {
                    .Text = TranslationManager.Instance.GetTranslation("NameColumn.HeaderText"),
                    .Width = 150
                                      },
                New ColumnHeader With {
                    .Text = TranslationManager.Instance.GetTranslation("DescriptionColumn.HeaderText"),
                    .Width = 200
                    },
                New ColumnHeader With {
                    .Text = TranslationManager.Instance.GetTranslation("NextRunColumn.HeaderText"),
                    .Width = 150
                    },
                New ColumnHeader With {
                    .Text = TranslationManager.Instance.GetTranslation("ScheduleTypeColumn.HeaderText"),
                    .Width = 100
                    },
                New ColumnHeader With {
                    .Text = TranslationManager.Instance.GetTranslation("ActionsColumn.HeaderText"),
                    .Width = 100
                    },
                New ColumnHeader With {
                    .Text = TranslationManager.Instance.GetTranslation("StatusColumn.HeaderText"),
                    .Width = 80
                    }
                })

                ' Add event handlers
                AddHandler _taskListView.SelectedIndexChanged, AddressOf TaskListView_SelectedIndexChanged
                AddHandler _taskListView.MouseClick, AddressOf TaskListView_MouseClick

                RefreshTaskList()
                _logger.LogInfo("Task list initialized successfully")
                AddHandler _taskListView.DoubleClick, AddressOf TaskListView_DoubleClick
            Catch ex As Exception
                _logger.LogError($"Error initializing task list: {ex.Message}")
                MessageBox.Show($"Error initializing task list: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub InstallModulesMenuItem_Click(sender As Object, e As EventArgs)
            Dim scriptPath As String = Path.Combine(Application.StartupPath, "LiteTaskData", "InstallModules.ps1")
            Dim startInfo As New ProcessStartInfo()
            startInfo.FileName = "powershell.exe"
            startInfo.Arguments = $"-NoProfile -ExecutionPolicy Bypass -File ""{scriptPath}"""
            startInfo.UseShellExecute = False
            startInfo.RedirectStandardOutput = True
            startInfo.RedirectStandardError = True
            startInfo.CreateNoWindow = True

            Dim process As New Process()
            process.StartInfo = startInfo
            process.Start()

            Dim output As String = process.StandardOutput.ReadToEnd()
            Dim [error] As String = process.StandardError.ReadToEnd()
            process.WaitForExit()

            If process.ExitCode = 0 Then
                MessageBox.Show("Modules installés avec succès.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show($"Erreur lors de l'installation des modules : {[error]}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Sub

        Private Function IsElevated() As Boolean
            Return New Security.Principal.WindowsPrincipal(Security.Principal.WindowsIdentity.GetCurrent()).IsInRole(Security.Principal.WindowsBuiltInRole.Administrator)
        End Function

        Private Sub LoadSettings()
            Try
                Me.Size = New Size(
            Integer.Parse(_xmlManager.ReadValue("MainForm", "Width", "800")),
            Integer.Parse(_xmlManager.ReadValue("MainForm", "Height", "600"))
        )

                _logger?.LogInfo("MainForm settings loaded successfully")
            Catch ex As Exception
                _logger?.LogError($"Error loading settings: {ex.Message}")
                MessageBox.Show($"Error loading settings: {ex.Message}",
                       "Settings Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Sub

        'Private Sub LogMessage(message As String)
        '    Dim logEntry = $"[{DateTime.Now}] {message}"
        '    Try
        '        ' Log to file
        '        Dim logPath As String = Path.Combine(Application.StartupPath, "app_log.txt")
        '        Using writer As New StreamWriter(logPath, True)
        '            writer.WriteLine(logEntry)
        '        End Using

        '        ' Log to GUI if in GUI mode
        '        If _isGuiMode Then
        '            If _logTextBox.InvokeRequired Then
        '                _logTextBox.Invoke(Sub() _logTextBox.AppendText(logEntry & Environment.NewLine))
        '            Else
        '                _logTextBox.AppendText(logEntry & Environment.NewLine)
        '            End If
        '        Else
        '            ' Log to console if not in GUI mode
        '            Console.WriteLine(logEntry)
        '        End If
        '    Catch ex As Exception
        '        ' If we can't log, try to show a message box in GUI mode, otherwise write to console
        '        If _isGuiMode Then
        '            MessageBox.Show($"Error logging message: {ex.Message}", "Logging Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        Else
        '            Console.WriteLine($"Error logging message: {ex.Message}")
        '        End If
        '    End Try
        'End Sub

        Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
            Try
                If _logger IsNot Nothing Then
                    _logger.LogInfo("MainForm closing")
                End If

                If _statusTimer IsNot Nothing Then
                    _statusTimer.Stop()
                    RemoveHandler _statusTimer.Tick, AddressOf UpdateStatus
                    _statusTimer.Dispose()
                End If

                If _schedulerTimer IsNot Nothing Then
                    _schedulerTimer.Stop()
                    RemoveHandler _schedulerTimer.Tick, AddressOf SchedulerTimer_Tick
                    _schedulerTimer.Dispose()
                End If

                'If _logger IsNot Nothing Then
                '    RemoveHandler _logger.LogEntryAdded, AddressOf OnLogEntryAdded
                'End If

                ' Save any pending changes
                If _customScheduler IsNot Nothing Then
                    _customScheduler.SaveTasks()
                End If

                MyBase.OnFormClosing(e)
            Catch ex As Exception
                If _logger IsNot Nothing Then
                    _logger.LogError("Error during form closing", ex)
                End If
            End Try
        End Sub

        'Private Sub OnLogEntryAdded(sender As Object, e As LogEntryEventArgs)
        '    Try
        '        If e.LogEntry.Level >= Logger.LogLevel.Warning Then
        '            UpdateStatusLabel(e.LogEntry.Message, If(e.LogEntry.Level = Logger.LogLevel.Error, Color.Red, Color.Orange))
        '        End If

        '        ' Add to log text box if visible
        '        If Not _splitContainer.Panel2Collapsed Then
        '            AppendToLog($"[{e.LogEntry.Level}] {e.LogEntry.Message}")
        '        End If

        '        ' Send notification for errors if enabled
        '        If e.LogEntry.Level >= Logger.LogLevel.Error Then
        '            _notificationManager.QueueNotification(
        '                $"LiteTask {e.LogEntry.Level}",
        '                e.LogEntry.Message,
        '                If(e.LogEntry.Level = Logger.LogLevel.Critical, NotificationPriority.High, NotificationPriority.Normal))
        '        End If
        '    Catch ex As Exception
        '        ' Use Debug.WriteLine as last resort since we can't use logger here
        '        Debug.WriteLine($"Error handling log entry: {ex.Message}")
        '    End Try
        'End Sub

        Private Sub OpenCredentialManager(sender As Object, e As EventArgs)
            If _credentialManager Is Nothing Then
                MessageBox.Show("Credential Manager is not initialized.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Using credentialManagerForm As New CredentialManagerForm
                credentialManagerForm.ShowDialog()
            End Using
        End Sub

        Private Async Sub RefreshAfterTaskExecution()
            Try
                Await Task.Delay(1000)
                Await RefreshUIAsync()
            Catch ex As Exception
                _logger?.LogError($"Error in RefreshAfterTaskExecution: {ex.Message}")
            End Try
        End Sub

        Private Async Function RefreshUIAsync() As Task
            Try
                _logger.LogInfo("Starting UI refresh")

                ' Reload tasks from storage
                _customScheduler.LoadTasks()

                ' Update ListView
                _taskListView.BeginUpdate()
                _taskListView.Items.Clear()

                Dim tasks = _customScheduler.GetAllTasks()
                If tasks IsNot Nothing Then
                    For Each task In tasks
                        If task IsNot Nothing Then
                            Dim item As New ListViewItem(task.Name) With {
                        .Tag = task  ' Store the entire task object in the Tag property
                    }
                            item.SubItems.Add(task.Description)
                            item.SubItems.Add(If(task.NextRunTime = DateTime.MaxValue,
                                      "One-time (Completed)",
                                      task.NextRunTime.ToString("g")))
                            item.SubItems.Add(task.Schedule.ToString())
                            item.SubItems.Add(task.Actions.Count.ToString() & " action(s)")

                            ' Set status and color
                            Dim status = If(task.NextRunTime > DateTime.Now, "Pending", "Due")
                            item.SubItems.Add(status)
                            If status = "Due" Then
                                item.BackColor = Color.LightYellow
                            End If

                            _taskListView.Items.Add(item)
                        End If
                    Next
                End If

                _taskListView.EndUpdate()
                _logger.LogInfo("UI refresh completed successfully")

            Catch ex As Exception
                _logger.LogError($"Error in RefreshUIAsync: {ex.Message}")
                _logger.LogError($"StackTrace: {ex.StackTrace}")
                MessageBox.Show($"Error refreshing UI: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Function

        Private Sub MonitorTasks_Click(sender As Object, e As EventArgs)
            Try
                ' Launch LitePM to show processes related to LiteTask
                _toolManager.LaunchLitePM("LiteTask")
            Catch ex As Exception
                _logger.LogError($"Error launching Lite Process Manager: {ex.Message}")
                _logger.LogError($"StackTrace: {ex.StackTrace}")
                MessageBox.Show($"Error launching Lite Process Manager: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        'Private Sub RunSelectedTasks(sender As Object, e As EventArgs)
        '    Try
        '        _logger.LogInfo("Run selected tasks menu item clicked")
        '        If _taskListView.SelectedItems.Count = 0 Then
        '            MessageBox.Show("Please select at least one task to run.", "No Task Selected", MessageBoxButtons.OK, MessageBoxIcon.Information)
        '            Return
        '        End If

        '        For Each item As ListViewItem In _taskListView.SelectedItems
        '            Dim taskName = item.Text
        '            Dim task = _customScheduler.GetTask(taskName)
        '            If task IsNot Nothing Then
        '                UpdateStatusLabel($"Running task: {task.Name}")
        '                _customScheduler.ExecuteTask(task).Wait()
        '                UpdateStatusLabel($"Task completed: {task.Name}")
        '            End If
        '        Next

        '        ' Refresh UI after all tasks complete
        '        RefreshTaskList()
        '        _refreshPending = True
        '        _autoRefreshTimer.Start()

        '    Catch ex As Exception
        '        _logger.LogError($"Error running selected tasks: {ex.Message}")
        '        UpdateStatusLabel("Error running tasks", Color.Red)
        '        MessageBox.Show($"Error running tasks: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    End Try
        'End Sub

        Private Async Sub RunSelectedTasks_Click(sender As Object, e As EventArgs)
            '_runButton.Enabled = False
            Try
                _logger.LogInfo("Run task button clicked")

                If _taskListView.SelectedItems.Count = 0 Then
                    MessageBox.Show("Please select a task to run.", "No Task Selected", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim selectedItem = _taskListView.SelectedItems(0)
                Dim taskName = selectedItem.Text
                Dim task = _customScheduler.GetTask(taskName)

                If task Is Nothing Then
                    _logger.LogError($"Task not found: {taskName}")
                    MessageBox.Show("Selected task could not be loaded.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                UpdateStatusLabel(String.Format(TranslationManager.Instance.GetTranslation("Status.RunningTask", "Running task: {0}"), task.Name))

                Try
                    Await _customScheduler.RunTaskAsync(task)
                    UpdateStatusLabel(String.Format(TranslationManager.Instance.GetTranslation("Status.TaskCompleted", "Task completed: {0}"), task.Name))

                    ' Safe UI update
                    If Not IsDisposed Then
                        If InvokeRequired Then
                            Invoke(Sub() RefreshTaskList())
                        Else
                            RefreshTaskList()
                        End If
                    End If

                Catch ex As Exception
                    _logger.LogError($"Error executing task: {ex.Message}")
                    UpdateStatusLabel(TranslationManager.Instance.GetTranslation("Status.ErrorRunningTask", "Error running task"), Color.Red)
                    MessageBox.Show($"Error running task: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Catch ex As Exception
                _logger.LogError($"Error in RunTaskButton_Click: {ex.Message}")
                MessageBox.Show($"Error running task: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                '_runButton.Enabled = True
            End Try
        End Sub

        Private Async Function RunTaskByNameAsync(taskName As String) As Task
            'TODO: Link to command line parameters to be able to launch tasks using their name.

            Try
                Dim task = _customScheduler.GetTask(taskName)
                If task IsNot Nothing Then
                    Await _customScheduler.RunTaskAsync(task)
                Else
                    _logger.LogError($"Task not found: {taskName}")
                    MessageBox.Show($"Task not found: {taskName}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Catch ex As Exception
                _logger.LogError($"Error running task {taskName}: {ex.Message}")
                MessageBox.Show($"Error running task {taskName}: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Function

        Private Sub RefreshMenuItem_Click(sender As Object, e As EventArgs)
            Try
                _logger.LogInfo("RefreshMenuItem clicked")
                If _customScheduler Is Nothing Then
                    _logger.LogError("CustomScheduler is null in RefreshMenuItem_Click")
                    MessageBox.Show("CustomScheduler is not initialized.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
                _customScheduler.LoadTasks()
                RefreshTaskList()
            Catch ex As Exception
                _logger.LogError($"Error in RefreshMenuItem_Click: {ex.Message}")
                _logger.LogError($"StackTrace: {ex.StackTrace}")
                MessageBox.Show($"An error occurred while refreshing: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub RefreshSettings()
            ' Update log level
            Dim logLevel As String = _xmlManager.ReadValue("Logging", "LogLevel", "Info")
            UpdateLogLevel(logLevel)

            ' Update log folder
            Dim logFolder As String = _xmlManager.ReadValue("Logging", "LogFolder", "")
            UpdateLogFolder(logFolder)

            ' Update alert settings
            Dim alertLevel As String = _xmlManager.ReadValue("Alerts", "AlertLevel", "Error")
            Dim alertEmail As String = _xmlManager.ReadValue("Alerts", "AlertEmail", "")
            UpdateAlertSettings(alertLevel, alertEmail)
        End Sub

        Private Sub TaskListView_MouseClick(sender As Object, e As MouseEventArgs) Handles _taskListView.MouseClick
            Try
                If e.Button = MouseButtons.Right Then
                    If Not VerifyDependencies() Then
                        Return
                    End If

                    ' Update menu item states based on selection
                    Dim hasSelection = _taskListView.SelectedItems.Count > 0
                    Dim singleSelection = _taskListView.SelectedItems.Count = 1

                    _runSelectedMenuItem.Enabled = hasSelection
                    _editSelectedMenuItem.Enabled = singleSelection
                    _deleteSelectedMenuItem.Enabled = hasSelection
                    _exportSelectedMenuItem.Enabled = hasSelection

                    ' Update menu item text to reflect selection count
                    If hasSelection AndAlso _taskListView.SelectedItems.Count > 1 Then
                        _runSelectedMenuItem.Text = $"Run Selected Tasks ({_taskListView.SelectedItems.Count})"
                        _deleteSelectedMenuItem.Text = $"Delete Selected Tasks ({_taskListView.SelectedItems.Count})"
                        _exportSelectedMenuItem.Text = $"Export Selected Tasks ({_taskListView.SelectedItems.Count})"
                    Else
                        _runSelectedMenuItem.Text = "Run Selected Task"
                        _deleteSelectedMenuItem.Text = "Delete Selected Task"
                        _exportSelectedMenuItem.Text = "Export Selected Task"
                    End If
                End If
            Catch ex As Exception
                _logger.LogError($"Error handling context menu display: {ex.Message}")
            End Try
        End Sub

        Private Sub RefreshTaskList()
            Try
                _logger.LogInfo("Starting task list refresh")
                _taskListView.BeginUpdate()
                _taskListView.Items.Clear()

                For Each task In _customScheduler.GetAllTasks()
                    If task IsNot Nothing Then
                        Dim item As New ListViewItem(task.Name) With {
                            .Text = task.Name
                        }
                        item.SubItems.Add(If(String.IsNullOrEmpty(task.Description), "-", task.Description))
                        item.SubItems.Add(If(task.NextRunTime = DateTime.MaxValue,
                                    "One-time (Completed)",
                                    task.NextRunTime.ToString("g")))
                        item.SubItems.Add(task.Schedule.ToString())
                        item.SubItems.Add($"{task.Actions.Count} action(s)")
                        item.SubItems.Add(GetTaskStatus(task))
                        _taskListView.Items.Add(item)
                    End If
                Next
                _taskListView.EndUpdate()
            Catch ex As Exception
                _logger.LogError($"Error refreshing task list: {ex.Message}")
                _logger.LogError($"StackTrace: {ex.StackTrace}")
                MessageBox.Show($"Error refreshing task list: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                _taskListView.EndUpdate()
            End Try
        End Sub

        Private Sub RequestRefresh()
            Try
                If IsDisposed OrElse Not Visible Then Return
                _refreshPending = True
                If _autoRefreshTimer IsNot Nothing Then
                    _autoRefreshTimer.Start()
                End If
            Catch ex As Exception
                _logger?.LogError($"Error in RequestRefresh: {ex.Message}")
            End Try
        End Sub

        Private Sub RestartAsAdmin(arguments As String)
            Try
                Dim startInfo As New ProcessStartInfo() With {
            .UseShellExecute = True,
            .WorkingDirectory = Environment.CurrentDirectory,
            .FileName = Application.ExecutablePath,
            .Verb = "runas"
        }

                If Not String.IsNullOrEmpty(arguments) Then
                    startInfo.Arguments = arguments
                End If

                Process.Start(startInfo)
                Application.Exit()
            Catch ex As Exception
                MessageBox.Show("Failed to restart with elevated privileges.",
                       "Elevation Error",
                       MessageBoxButtons.OK,
                       MessageBoxIcon.Error)
            End Try
        End Sub

        'Private Sub ScriptErrorReceived(sender As Object, e As DataAddedEventArgs)
        '    Dim psDataCollection = CType(sender, PSDataCollection(Of ErrorRecord))
        '    Dim data = psDataCollection(e.Index)
        '    LogMessage($"Error: {data.Exception.Message}")
        'End Sub

        'Private Sub ScriptOutputReceived(sender As Object, e As DataAddedEventArgs)
        '    Dim psDataCollection = CType(sender, PSDataCollection(Of InformationRecord))
        '    Dim data = psDataCollection(e.Index)
        '    LogMessage($"Output: {data.MessageData}")
        'End Sub

        'Public Sub SetTaskListView(taskListView As ListView)
        '    If taskListView Is Nothing Then
        '        Throw New ArgumentNullException(NameOf(taskListView))
        '    End If
        '    _taskListView = taskListView
        'End Sub

        Private Sub SettingsMenuItem_Click(sender As Object, e As EventArgs)
            Using optionsForm As New OptionsForm(_xmlManager)
                If optionsForm.ShowDialog() = DialogResult.OK Then
                    RefreshSettings()
                End If
            End Using
        End Sub

        Private Sub ShowHelp()
        End Sub

        Private Sub TaskListView_DoubleClick(sender As Object, e As EventArgs)
            Try
                If _taskListView.SelectedItems.Count = 1 Then
                    EditSelectedTasks_Click(sender, e)
                End If
            Catch ex As Exception
                _logger.LogError($"Error in double-click handler: {ex.Message}")
            End Try
        End Sub

        Private Sub TaskListView_SelectedIndexChanged(sender As Object, e As EventArgs)
            Try
                _editTaskButton.Enabled = (_taskListView.SelectedItems.Count = 1)
                _deleteTaskButton.Enabled = (_taskListView.SelectedItems.Count > 0)
                _runTaskButton.Enabled = (_taskListView.SelectedItems.Count > 0)
            Catch ex As Exception
                _logger.LogError($"Error in TaskListView_SelectedIndexChanged: {ex.Message}")
            End Try
        End Sub

        'Private Sub ToggleLogVisibility(sender As Object, e As EventArgs)
        '    Try
        '        If _splitContainer Is Nothing Then
        '            MessageBox.Show("Split container is not initialized.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '            Return
        '        End If

        '        _splitContainer.Panel2Collapsed = Not _splitContainer.Panel2Collapsed

        '        ' Update the menu item text
        '        If _toggleLogMenuItem IsNot Nothing Then
        '            _toggleLogMenuItem.Text = If(_splitContainer.Panel2Collapsed, "Show Log", "Hide Log")
        '        End If

        '        ' Ensure the form layout is updated
        '        Me.PerformLayout()
        '    Catch ex As Exception
        '        _logger?.LogError($"Error in ToggleLogVisibility: {ex.Message}")
        '        _logger?.LogError($"StackTrace: {ex.StackTrace}")
        '        MessageBox.Show($"An error occurred while toggling log visibility: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    End Try
        'End Sub

        Private Sub UpdateAlertSettings(alertLevel As String, alertEmail As String)
            ' TODO: Implement alert settings
            ' For example:
            ' _alertManager.SetAlertLevel(alertLevel)
            ' _alertManager.SetAlertEmail(alertEmail)
        End Sub

        'Private Sub UpdateLog(message As String)
        '    If _logTextBox.InvokeRequired Then
        '        _logTextBox.Invoke(Sub() UpdateLog(message))
        '    Else
        '        _logTextBox.AppendText(message & Environment.NewLine)
        '    End If
        'End Sub

        Private Sub UpdateStatus(sender As Object, e As EventArgs)
            Try
                Dim tasks = _customScheduler.GetAllTasks()
                Dim pendingTasks = tasks.Count(Function(t) t.NextRunTime > DateTime.Now)
                Dim dueTasks = tasks.Count(Function(t) t.NextRunTime <= DateTime.Now)

                Dim serviceNote = If(IsServiceRunning(), $" ({TranslationManager.Instance.GetTranslation("Status.ServiceActive", "service active")})", "")
                UpdateStatusLabel(String.Format(TranslationManager.Instance.GetTranslation("Status.Ready", "Ready - {0} pending tasks, {1} due tasks"), pendingTasks, dueTasks) & serviceNote)

            Catch ex As Exception
                _logger.LogError("Error updating status", ex)
                UpdateStatusLabel(TranslationManager.Instance.GetTranslation("Status.ErrorUpdating", "Error updating status"), Color.Red)
            End Try
        End Sub

        Private Sub SchedulerTimer_Tick(sender As Object, e As EventArgs)
            Try
                ' Skip GUI scheduling when the Windows service is already handling it
                If IsServiceRunning() Then
                    Return
                End If

                _customScheduler.CheckAndExecuteTasks()
            Catch ex As Exception
                _logger.LogError($"Error in scheduler timer: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Checks whether the LiteTask Windows service is currently running.
        ''' Result is cached to avoid querying the service controller too frequently.
        ''' </summary>
        Private Function IsServiceRunning() As Boolean
            Try
                Dim now = DateTime.Now
                If (now - _lastServiceCheck).TotalSeconds < SERVICE_CHECK_INTERVAL_SECONDS Then
                    Return _isServiceRunning
                End If

                _lastServiceCheck = now
                Using sc = New ServiceController(SERVICE_NAME)
                    _isServiceRunning = (sc.Status = ServiceControllerStatus.Running)
                End Using
            Catch
                ' Service not installed or not accessible — treat as not running
                _isServiceRunning = False
            End Try

            Return _isServiceRunning
        End Function

        Private Sub UpdateTaskStatus(task As ScheduledTask)
            'TODO: Implement task status update
            Try
                For Each item In _taskListView.Items.Cast(Of ListViewItem)()
                    If item.Text = task.Name Then
                        ' Update status based on actions
                        Dim status = GetTaskStatus(task)
                        item.SubItems("Status").Text = status

                        ' Update visual indicators
                        Select Case status
                            Case "Running"
                                item.BackColor = Color.LightBlue
                            Case "Failed"
                                item.BackColor = Color.LightPink
                            Case "Completed"
                                item.BackColor = Color.LightGreen
                            Case Else
                                item.BackColor = SystemColors.Window
                        End Select
                        Exit For
                    End If
                Next
            Catch ex As Exception
                _logger.LogError($"Error updating task status: {ex.Message}")
            End Try
        End Sub

        Private Sub UpdateLogFolder(logFolder As String)
            ' Update the logger's log folder
            If Not String.IsNullOrEmpty(logFolder) Then
                _logger.SetLogFolder(logFolder)
            End If
        End Sub

        Private Sub UpdateLogLevel(logLevel As String)
            ' Update the logger's log level
            Select Case logLevel.ToLower()
                Case "debug"
                    _logger.SetLogLevel(Logger.LogLevel.Debug)
                Case "info"
                    _logger.SetLogLevel(Logger.LogLevel.Info)
                Case "warning"
                    _logger.SetLogLevel(Logger.LogLevel.Warning)
                Case "error"
                    _logger.SetLogLevel(Logger.LogLevel.Error)
                Case "critical"
                    _logger.SetLogLevel(Logger.LogLevel.Critical)
            End Select
        End Sub

        Private Sub UpdateStatusLabel(message As String, Optional color As Color = Nothing)
            If Me.InvokeRequired Then
                Me.Invoke(Sub() UpdateStatusLabel(message, color))
            Else
                _statusLabel.Text = message
                If color <> Nothing Then
                    _statusLabel.ForeColor = color
                Else
                    _statusLabel.ForeColor = SystemColors.ControlText
                End If
            End If
        End Sub

        Private Async Sub UpdateTools_Click(sender As Object, e As EventArgs)
            _updateToolsMenuItem.Enabled = False
            Try
                Dim updateResult = Await _toolManager.DownloadAndUpdateAllToolsAsync()
                Dim message As String = "Update Result:" & Environment.NewLine
                For Each kvp In updateResult
                    message &= $"{kvp.Key}: {If(kvp.Value, "Updated", "Failed")}" & Environment.NewLine
                Next
                MessageBox.Show(message, "Tool Update Result", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show($"Error updating tools: {ex.Message}", "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                _updateToolsMenuItem.Enabled = True
            End Try
        End Sub

        Private Function VerifyDependencies() As Boolean
            Try
                Dim missingDependencies = New List(Of String)

                If _customScheduler Is Nothing Then
                    missingDependencies.Add("CustomScheduler")
                End If

                If _xmlManager Is Nothing Then
                    missingDependencies.Add("XMLManager")
                End If

                If _logger Is Nothing Then
                    missingDependencies.Add("Logger")
                End If

                If _taskListView Is Nothing Then
                    missingDependencies.Add("TaskListView")
                End If

                If missingDependencies.Count > 0 Then
                    Dim message = $"The following dependencies are not initialized: {String.Join(", ", missingDependencies)}"
                    _logger?.LogError(message)
                    MessageBox.Show(message, "Initialization Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return False
                End If

                Return True
            Catch ex As Exception
                _logger?.LogError($"Error verifying dependencies: {ex.Message}")
                MessageBox.Show($"Error verifying dependencies: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        End Function

    End Class
End Namespace