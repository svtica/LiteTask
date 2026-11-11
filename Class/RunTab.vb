Imports System.Data
Imports LiteTask.LiteTask.ScheduledTask

Namespace LiteTask
    Public Class RunTab
        Implements IDisposable
        Private ReadOnly _credentialManager As CredentialManager
        Private ReadOnly _customScheduler As CustomScheduler
        Private ReadOnly _logger As Logger
        Private ReadOnly _xmlManager As XMLManager
        Private ReadOnly _taskrunner As TaskRunner
        Private _disposed As Boolean = False
        Private ReadOnly _powerShellPathManager As PowerShellPathManager
        Private ReadOnly _translationManager As TranslationManager
        Private _targetBrowseDialog As OpenFileDialog
        Private _scriptTextBox As TextBox
        Private _outputTextBox As TextBox
        Private _runButton As Button
        Private _credentialComboBox As ComboBox
        Private _tabPage As TabPage
        Private _requiresElevationCheckBox As CheckBox
        Private _executionModeComboBox As ComboBox
        Private _executionModeLabel As Label
        Private _scriptLabel As Label
        Private _outputLabel As Label
        Private _credentialLabel As Label
        Private _targetTextBox As TextBox
        Private _targetLabel As Label
        Private _executionTypeLabel As Label
        Private _executionTypeComboBox As ComboBox
        Private _autoDetectCheckBox As CheckBox
        Private _tableLayoutPanel As TableLayoutPanel
        Private _tabPageInitialized As Boolean = False
        Public Event OutputDataReceived As EventHandler(Of DataAddedEventArgs)
        Public Event ErrorDataReceived As EventHandler(Of DataAddedEventArgs)

        Public Sub New(credentialManager As CredentialManager, logger As Logger, taskRunner As TaskRunner, xmlManager As XMLManager, customScheduler As CustomScheduler)
            Try
                _credentialManager = credentialManager
                _logger = logger
                _xmlManager = xmlManager
                _taskrunner = taskRunner
                _customScheduler = customScheduler
                _powerShellPathManager = New PowerShellPathManager(logger)

                ' Initialize TranslationManager
                _translationManager = TranslationManager.Initialize(logger, xmlManager)
                If _translationManager Is Nothing Then
                    _logger?.LogError("Failed to initialize TranslationManager")
                End If

                _logger.LogInfo("RunTab initialized successfully")

            Catch ex As Exception
                _logger?.LogError($"Error in RunTab constructor: {ex.Message}")
                _logger?.LogError($"Stack trace: {ex.StackTrace}")
                Throw
            End Try
        End Sub

        Private Sub HandleError(sender As Object, data As String)
            If _outputTextBox.IsHandleCreated Then
                _outputTextBox.Invoke(Sub()
                                          _outputTextBox.AppendText("Error: " & data & Environment.NewLine)
                                          _outputTextBox.ScrollToCaret()
                                      End Sub)
            End If
        End Sub

        Private Sub HandleOutput(sender As Object, data As String)
            If _outputTextBox.IsHandleCreated Then
                _outputTextBox.Invoke(Sub()
                                          _outputTextBox.AppendText(data & Environment.NewLine)
                                          _outputTextBox.ScrollToCaret()
                                      End Sub)
            End If
        End Sub

        Private Sub InitializeComponent()
            Try
                _logger?.LogInfo("Starting InitializeComponent for RunTab")

                ' Initialize TabPage
                _tabPage = New TabPage("Run")

                ' Initialize Labels
                _scriptLabel = New Label With {
            .Text = "Script:",
            .Dock = DockStyle.Fill,
            .AutoSize = True
        }

                _outputLabel = New Label With {
            .Text = "Output:",
            .Dock = DockStyle.Fill,
            .AutoSize = True
        }

                _targetLabel = New Label With {
            .Text = "Target:",
            .Dock = DockStyle.Fill,
            .AutoSize = True
        }

                _credentialLabel = New Label With {
            .Text = "Credential:",
            .Dock = DockStyle.Fill,
            .AutoSize = True
        }

                _executionTypeLabel = New Label With {
            .Text = "Execution Type:",
            .Dock = DockStyle.Fill,
            .AutoSize = True
        }

                ' Initialize other controls
                _tableLayoutPanel = New TableLayoutPanel()
                _tableLayoutPanel.Dock = DockStyle.Fill
                _tableLayoutPanel.ColumnCount = 2
                _tableLayoutPanel.RowCount = 8
                _tableLayoutPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 20))
                _tableLayoutPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 80))

                ' Add rows with appropriate sizing
                _tableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 35))  ' Script
                _tableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 35))  ' Output
                _tableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 30)) ' Target
                _tableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 30)) ' Execution Type
                _tableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 30)) ' Credential
                _tableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 30)) ' Requires Elevation
                _tableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 40)) ' Run Button


                _scriptTextBox = New TextBox With {
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical,
            .Dock = DockStyle.Fill
        }

                _outputTextBox = New TextBox With {
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical,
            .Dock = DockStyle.Fill,
            .ReadOnly = True
        }
                _targetTextBox = New TextBox With {
    .Dock = DockStyle.Fill
}

                _runButton = New Button With {
            .Text = "Run Script",
            .Dock = DockStyle.Fill,
            .Height = 30
        }

                _credentialComboBox = New ComboBox With {
            .Dock = DockStyle.Fill,
            .DropDownStyle = ComboBoxStyle.DropDownList
        }

                _requiresElevationCheckBox = New CheckBox With {
            .Text = "Run with elevated privileges",
            .Dock = DockStyle.Fill
        }

                _autoDetectCheckBox = New CheckBox With {
            .Text = "Auto-detect execution type",
            .Checked = True,
            .Dock = DockStyle.Fill
        }

                _executionModeComboBox = New ComboBox With {
            .Dock = DockStyle.Fill,
            .DropDownStyle = ComboBoxStyle.DropDownList
        }

                _executionTypeComboBox = New ComboBox With {
            .Dock = DockStyle.Fill,
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
                ' Only real TaskType values ("Command" used to be listed and made Enum.Parse throw)
                _executionTypeComboBox.Items.AddRange({TaskType.PowerShell.ToString(), TaskType.Batch.ToString(), TaskType.SQL.ToString(), TaskType.Executable.ToString()})
                _executionTypeComboBox.SelectedIndex = 0

                _autoDetectCheckBox = New CheckBox With {
                    .Text = "Auto-detect execution type",
                    .Checked = True,
                    .Dock = DockStyle.Fill
                }

                ' Add translations with null checks
                Try
                    _logger?.LogInfo("Starting translations for RunTab")
                    If _translationManager IsNot Nothing Then
                        _tabPage.Text = _translationManager.GetTranslation("RunTab.Text")
                        _scriptLabel.Text = _translationManager.GetTranslation("RunTab.ScriptLabel")
                        _outputLabel.Text = _translationManager.GetTranslation("RunTab.OutputLabel")
                        _targetLabel.Text = _translationManager.GetTranslation("RunTab.TargetLabel")
                        _credentialLabel.Text = _translationManager.GetTranslation("RunTab.CredentialLabel")
                        _executionTypeLabel.Text = _translationManager.GetTranslation("RunTab.ExecutionTypeLabel")
                        _requiresElevationCheckBox.Text = _translationManager.GetTranslation("RunTab.RequiresElevation")
                        _autoDetectCheckBox.Text = _translationManager.GetTranslation("RunTab.AutoDetect")
                        _runButton.Text = _translationManager.GetTranslation("RunTab.RunButton")
                        _logger?.LogInfo("RunTab translations completed")
                    End If
                Catch translationEx As Exception
                    _logger?.LogError($"Error applying translations: {translationEx.Message}")
                    _logger?.LogError($"Stack trace: {translationEx.StackTrace}")
                End Try

                ' Add controls to TableLayoutPanel
                _tableLayoutPanel.Controls.Add(_scriptLabel, 0, 0)
                _tableLayoutPanel.Controls.Add(_scriptTextBox, 1, 0)
                _tableLayoutPanel.Controls.Add(_outputLabel, 0, 1)
                _tableLayoutPanel.Controls.Add(_outputTextBox, 1, 1)
                _tableLayoutPanel.Controls.Add(_targetLabel, 0, 2)
                _tableLayoutPanel.Controls.Add(_targetTextBox, 1, 2)
                _tableLayoutPanel.Controls.Add(_executionTypeLabel, 0, 3)
                _tableLayoutPanel.Controls.Add(_executionTypeComboBox, 1, 3)
                _tableLayoutPanel.Controls.Add(_credentialLabel, 0, 4)
                _tableLayoutPanel.Controls.Add(_credentialComboBox, 1, 4)
                _tableLayoutPanel.Controls.Add(_requiresElevationCheckBox, 1, 5)
                _tableLayoutPanel.Controls.Add(_runButton, 1, 6)

                ' Add TableLayoutPanel to TabPage
                _tabPage.Controls.Add(_tableLayoutPanel)
                AddHandlers()
                PopulateQueryTypeComboBox()

                ' TaskRunner output is subscribed per run in RunButton_Click (filtered to this
                ' tab's run): a permanent subscription also showed every scheduled task's output.

                _logger?.LogInfo("InitializeComponent completed successfully for RunTab")

            Catch ex As Exception
                _logger?.LogError($"Error in InitializeComponent: {ex.Message}")
                _logger?.LogError($"StackTrace: {ex.StackTrace}")
                Throw
            End Try
        End Sub

        Public Sub AddHandlers()
            AddHandler _executionTypeComboBox.SelectedIndexChanged, AddressOf ExecuteTypeComboBox_SelectedIndexChanged
            AddHandler _autoDetectCheckBox.CheckedChanged, AddressOf AutoDetectCheckBox_CheckedChanged
            AddHandler _runButton.Click, AddressOf RunButton_Click
        End Sub

        Private Sub AutoDetectCheckBox_CheckedChanged(sender As Object, e As EventArgs)
            ' Auto-detect only applies to a target file; the type list stays usable for typed commands
        End Sub

        Private Sub AutoDetectExecutionType()
            Try
                If _autoDetectCheckBox.Checked AndAlso Not String.IsNullOrEmpty(_targetTextBox.Text) Then
                    Dim detectedType = FileTypeDetector.DetectTaskType(_targetTextBox.Text)
                    _executionTypeComboBox.SelectedItem = detectedType.ToString()
                End If
            Catch ex As Exception
                _logger.LogError($"Error in AutoDetectExecutionType: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Builds the action to run. With a target file, the Script box holds its
        ''' parameters (as before). Without one, the Script box IS the script: it is
        ''' written to a temporary .ps1/.bat/.sql file (TaskRunner executes files), or
        ''' passed to cmd /c for the Executable type. Typed commands used to be sent as
        ''' parameters of a missing file, which made every run fail.
        ''' </summary>
        Private Function CreateTaskAction(ByRef tempScriptPath As String) As TaskAction
            tempScriptPath = Nothing
            Dim actionType = GetExecutionType()
            Dim target = If(String.IsNullOrWhiteSpace(_targetTextBox.Text), Nothing, _targetTextBox.Text.Trim())
            Dim parameters As String = _scriptTextBox.Text

            If target Is Nothing AndAlso actionType <> TaskType.Executable Then
                Dim extension = ScriptExtension(actionType)
                Dim tempDir = Path.Combine(Application.StartupPath, "LiteTaskData", "temp")
                Directory.CreateDirectory(tempDir)
                tempScriptPath = Path.Combine(tempDir, $"runtab_{DateTime.Now:yyyyMMddHHmmss}_{Guid.NewGuid():N}{extension}")
                ' Batch files are read by cmd.exe with the OEM code page; PowerShell and sqlcmd read UTF-8 with BOM
                If actionType = TaskType.Batch Then
                    File.WriteAllText(tempScriptPath, _scriptTextBox.Text, Text.Encoding.GetEncoding(Globalization.CultureInfo.CurrentCulture.TextInfo.OEMCodePage))
                Else
                    File.WriteAllText(tempScriptPath, _scriptTextBox.Text, New Text.UTF8Encoding(True))
                End If
                target = tempScriptPath
                parameters = Nothing
            End If

            Return New TaskAction With {
                .Name = $"RunTab_{DateTime.Now:yyyyMMddHHmmss}",
                .Type = actionType,
                .Target = target,
                .Parameters = parameters,
                .RequiresElevation = _requiresElevationCheckBox.Checked
            }
        End Function

        Private Shared Function ScriptExtension(actionType As TaskType) As String
            Select Case actionType
                Case TaskType.Batch : Return ".bat"
                Case TaskType.SQL : Return ".sql"
                Case Else : Return ".ps1"
            End Select
        End Function

        Private Function DataTableToString(dt As DataTable) As String
            Dim result As New StringBuilder()

            ' Add headers
            result.AppendLine(String.Join(vbTab, dt.Columns.Cast(Of DataColumn)().Select(Function(c) c.ColumnName)))

            ' Add rows
            For Each row As DataRow In dt.Rows
                result.AppendLine(String.Join(vbTab, row.ItemArray.Select(Function(item) If(item IsNot Nothing, item.ToString(), ""))))
            Next

            Return result.ToString()
        End Function
        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overrides Sub Finalize()
            Dispose(False)
            MyBase.Finalize()
        End Sub

        Private Sub ThrowIfDisposed()
            If _disposed Then
                Throw New ObjectDisposedException(GetType(RunTab).FullName)
            End If
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not _disposed Then
                If disposing Then
                    ' Clean up managed resources
                    RemoveHandler _taskrunner.OutputReceived, AddressOf HandleOutput
                    RemoveHandler _taskrunner.ErrorReceived, AddressOf HandleError

                    If _targetBrowseDialog IsNot Nothing Then
                        _targetBrowseDialog.Dispose()
                        _targetBrowseDialog = Nothing
                    End If
                End If

                ' Clean up unmanaged resources (if any) here

                _disposed = True
            End If
        End Sub

        Private Sub ErrorHandler(sender As Object, e As DataAddedEventArgs)
            If _outputTextBox.IsHandleCreated Then
                _outputTextBox.Invoke(Sub()
                                          Dim err = CType(sender, PSDataCollection(Of ErrorRecord))(e.Index).ToString()
                                          _outputTextBox.AppendText("Error: " & err & Environment.NewLine)
                                      End Sub)
            End If
        End Sub

        Private Async Function ExecuteAction(action As TaskAction, credential As CredentialInfo) As Task(Of Boolean)
            _logger.LogInfo($"Executing {action.Type} action: {action.Name}")
            _outputTextBox.Clear()

            Try
                Select Case action.Type
                    Case TaskType.SQL
                        Return Await _taskrunner.ExecuteSqlTask(action, credential)
                    Case TaskType.PowerShell
                        Return Await _taskrunner.ExecutePowerShellTask(action, credential)
                    Case TaskType.Batch
                        Return Await _taskrunner.ExecuteBatchTask(action, credential)
                    Case TaskType.Executable
                        Return Await _taskrunner.ExecuteExecutableTask(action, credential)
                    Case Else
                        Throw New NotSupportedException($"Task type {action.Type} is not supported")
                End Select
            Finally
                ' Reclaim memory after the run (Reliability/ForceGcAfterTask)
                MemoryDiagnostics.CollectAfterTask()
            End Try
        End Function

        Private Sub ExecuteTypeComboBox_SelectedIndexChanged(sender As Object, e As EventArgs)
            Try
                If _targetBrowseDialog Is Nothing Then
                    _targetBrowseDialog = New OpenFileDialog()
                End If

                If _executionTypeComboBox.SelectedItem.ToString().Equals("SQL", StringComparison.OrdinalIgnoreCase) Then
                    _targetBrowseDialog.Filter = "SQL files (*.sql)|*.sql|All files (*.*)|*.*"
                Else
                    _targetBrowseDialog.Filter = "All files (*.*)|*.*"
                End If
            Catch ex As Exception
                _logger.LogError($"Error in ExecuteTypeComboBox_SelectedIndexChanged: {ex.Message}")
            End Try
        End Sub

        Public Function GetTabPage() As TabPage
            If Not _tabPageInitialized Then
                Try
                    _logger?.LogInfo("Initializing RunTab components")
                    InitializeComponent()
                    PopulateCredentialComboBox()
                    _tabPageInitialized = True
                    _logger?.LogInfo("RunTab components initialized successfully")
                Catch ex As Exception
                    _logger?.LogError($"Error initializing RunTab components: {ex.Message}")
                    _logger?.LogError($"StackTrace: {ex.StackTrace}")
                    Throw
                End Try
            End If
            Return _tabPage
        End Function

        Protected Sub OnShown(e As EventArgs)
            OnShown(e)
            AutoDetectExecutionType()
        End Sub

        Private Sub OutputHandler(sender As Object, e As DataAddedEventArgs)
            If _outputTextBox.IsHandleCreated Then
                _outputTextBox.Invoke(Sub()
                                          Dim output = CType(sender, PSDataCollection(Of InformationRecord))(e.Index).MessageData.ToString()
                                          _outputTextBox.AppendText(output & Environment.NewLine)
                                      End Sub)
            End If
        End Sub

        Private Sub PopulateCredentialComboBox()
            _credentialComboBox.Items.Clear()
            _credentialComboBox.Items.Add("(None)")
            Dim targets = _credentialManager.GetAllCredentialTargets()
            For Each target In targets
                _credentialComboBox.Items.Add(target)
            Next
            _credentialComboBox.SelectedIndex = 0
        End Sub

        Private Sub PopulateQueryTypeComboBox()
            _executionTypeComboBox.Items.Clear()
            _executionTypeComboBox.Items.AddRange({
        TaskType.PowerShell.ToString(),
        TaskType.Batch.ToString(),
        TaskType.SQL.ToString(),
        TaskType.Executable.ToString()
    })
            _executionTypeComboBox.SelectedIndex = 0
        End Sub

        Private Sub RaiseErrorEvent(sender As Object, e As DataAddedEventArgs)
            RaiseEvent ErrorDataReceived(sender, e)
            _logger.LogError($"Process error: {e.ToString()}")
        End Sub

        Private Sub RaiseOutputEvent(sender As Object, e As DataAddedEventArgs)
            RaiseEvent OutputDataReceived(sender, e)
            _logger.LogInfo($"Process output: {e.ToString()}")
        End Sub

        Private Async Sub RunButton_Click(sender As Object, e As EventArgs)
            Dim tempScriptPath As String = Nothing
            ' Show the run's output in the Output box (only events from this ad hoc run:
            ' scheduled tasks running at the same time carry their task name)
            Dim onOutput As TaskRunner.OutputReceivedEventHandler = Sub(s, data) If String.IsNullOrEmpty(_taskrunner.CurrentTaskName) Then HandleOutput(s, data)
            Dim onError As TaskRunner.ErrorReceivedEventHandler = Sub(s, data) If String.IsNullOrEmpty(_taskrunner.CurrentTaskName) Then HandleError(s, data)
            Try
                _logger.LogInfo("Starting execution")
                _runButton.Enabled = False
                _outputTextBox.Clear()

                If Not ValidateInputs() Then Return

                Dim selectedCredential = If(_credentialComboBox?.SelectedItem?.ToString(), "(None)")
                Dim credential = If(selectedCredential = "(None)",
                              Nothing,
                              _credentialManager.GetCredential(selectedCredential, "Windows Vault"))

                Dim action = CreateTaskAction(tempScriptPath)
                AddHandler _taskrunner.OutputReceived, onOutput
                AddHandler _taskrunner.ErrorReceived, onError
                Dim success = Await ExecuteAction(action, credential)

                If success Then
                    _outputTextBox.AppendText(Environment.NewLine & "Task executed successfully" & Environment.NewLine)
                Else
                    _outputTextBox.AppendText(Environment.NewLine & "Task execution failed" & Environment.NewLine)
                End If

            Catch ex As Exception
                _logger.LogError($"Error in RunButton_Click: {ex.Message}")
                _outputTextBox.AppendText($"Error executing task: {ex.Message}" & Environment.NewLine)
            Finally
                RemoveHandler _taskrunner.OutputReceived, onOutput
                RemoveHandler _taskrunner.ErrorReceived, onError
                If tempScriptPath IsNot Nothing Then
                    Try
                        File.Delete(tempScriptPath)
                    Catch
                    End Try
                End If
                _runButton.Enabled = True
            End Try
        End Sub

        Private Function GetExecutionType() As TaskType
            ' Auto-detect needs a file name; typed commands use the selected type
            If _autoDetectCheckBox.Checked AndAlso Not String.IsNullOrWhiteSpace(_targetTextBox.Text) Then
                Return FileTypeDetector.DetectTaskType(_targetTextBox.Text)
            End If
            Dim selected = If(_executionTypeComboBox.SelectedItem?.ToString(), TaskType.PowerShell.ToString())
            Dim parsed As TaskType
            Return If([Enum].TryParse(selected, parsed), parsed, TaskType.PowerShell)
        End Function

        Private Function ValidateInputs() As Boolean
            Dim hasTarget = Not String.IsNullOrWhiteSpace(_targetTextBox.Text)
            If Not hasTarget AndAlso String.IsNullOrWhiteSpace(_scriptTextBox.Text) Then
                MessageBox.Show(TranslationManager.Instance.GetTranslation("RunTab.NothingToRun", "Type a script or command, or choose a target file."),
                                "LiteTask", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return False
            End If
            If hasTarget AndAlso Not File.Exists(_targetTextBox.Text.Trim()) Then
                MessageBox.Show(String.Format(TranslationManager.Instance.GetTranslation("RunTab.TargetNotFound", "Target file not found: {0}"), _targetTextBox.Text.Trim()),
                                "LiteTask", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If
            Return True
        End Function

    End Class
End Namespace