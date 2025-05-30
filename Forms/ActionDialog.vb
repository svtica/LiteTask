Imports LiteTask.LiteTask.ScheduledTask

Namespace LiteTask
    Partial Public Class ActionDialog
        Inherits Form

        Private ReadOnly _logger As Logger
        Private _existingAction As TaskAction
        Private _isEditMode As Boolean

        ' UI Controls
        Friend WithEvents _nameLabel As Label
        Friend WithEvents _nameTextBox As TextBox
        Friend WithEvents _typeLabel As Label
        Friend WithEvents _typeComboBox As ComboBox
        Friend WithEvents _targetLabel As Label
        Friend WithEvents _targetTextBox As TextBox
        Friend WithEvents _browseButton As Button
        Friend WithEvents _parametersLabel As Label
        Friend WithEvents _parametersTextBox As TextBox
        Friend WithEvents _dependsOnLabel As Label
        Friend WithEvents _dependsOnCombo As ComboBox
        Friend WithEvents _waitForCompletionCheck As CheckBox
        Friend WithEvents _timeoutLabel As Label
        Friend WithEvents _timeoutNumeric As NumericUpDown
        Friend WithEvents _retryCountLabel As Label
        Friend WithEvents _retryCountNumeric As NumericUpDown
        Friend WithEvents _retryDelayLabel As Label
        Friend WithEvents _retryDelayNumeric As NumericUpDown
        Friend WithEvents _continueOnErrorCheck As CheckBox
        Friend WithEvents _okButton As Button
        Friend WithEvents _cancelButton As Button
        Friend WithEvents _requiresElevationCheckBox As CheckBox

        Public Property Action As TaskAction

        Public Sub New(Optional existingAction As TaskAction = Nothing)
            Try
                ' Store dependencies
                _logger = ApplicationContainer.GetService(Of Logger)()
                _existingAction = existingAction
                _isEditMode = (existingAction IsNot Nothing)

                ' Initialize form components
                InitializeComponent()

                ' Initialize ComboBoxes before loading data
                InitializeComboBoxes()

                ' Now load the data
                InitializeData()
                AddEventsHandlers()
                Me.Translate()

            Catch ex As Exception
                _logger?.LogError($"Error initializing ActionDialog: {ex.Message}")
                Throw New InvalidOperationException("Failed to initialize action dialog", ex)
            End Try
        End Sub

        Private Sub InitializeComboBoxes()
            Try
                ' Initialize Type ComboBox
                _typeComboBox.Items.Clear()
                _typeComboBox.Items.AddRange([Enum].GetNames(GetType(TaskType)))
                _typeComboBox.SelectedIndex = 0

                ' Initialize Depends On ComboBox
                _dependsOnCombo.Items.Clear()
                _dependsOnCombo.Items.Add("(None)")
                _dependsOnCombo.SelectedIndex = 0

                PopulateDependencyCombo()

            Catch ex As Exception
                _logger?.LogError($"Error initializing ComboBoxes: {ex.Message}")
                Throw
            End Try
        End Sub

        Private Sub InitializeData()
            Try
                If _existingAction IsNot Nothing Then
                    _nameTextBox.Text = _existingAction.Name

                    ' Set Type ComboBox
                    Dim typeIndex = _typeComboBox.Items.IndexOf(_existingAction.Type.ToString())
                    If typeIndex >= 0 Then
                        _typeComboBox.SelectedIndex = typeIndex
                    End If

                    _targetTextBox.Text = _existingAction.Target
                    _parametersTextBox.Text = _existingAction.Parameters
                    _requiresElevationCheckBox.Checked = _existingAction.RequiresElevation

                    ' Load dependency settings
                    If _existingAction.DependsOn IsNot Nothing Then
                        Dim index = _dependsOnCombo.Items.IndexOf(_existingAction.DependsOn)
                        If index >= 0 Then
                            _dependsOnCombo.SelectedIndex = index
                        End If
                    End If

                    _waitForCompletionCheck.Checked = _existingAction.WaitForCompletion
                    _timeoutNumeric.Value = _existingAction.TimeoutMinutes
                    _retryCountNumeric.Value = _existingAction.RetryCount
                    _retryDelayNumeric.Value = _existingAction.RetryDelayMinutes
                    _continueOnErrorCheck.Checked = _existingAction.ContinueOnError
                End If

                UpdateDependencyControls(Nothing, EventArgs.Empty)

            Catch ex As Exception
                _logger?.LogError($"Error initializing data: {ex.Message}")
                Throw
            End Try
        End Sub

        Private Sub AddEventsHandlers()
            AddHandler _typeComboBox.SelectedIndexChanged, AddressOf TypeComboBox_SelectedIndexChanged
            AddHandler _browseButton.Click, AddressOf BrowseButton_Click
            AddHandler _okButton.Click, AddressOf OkButton_Click
            AddHandler _waitForCompletionCheck.CheckedChanged, AddressOf UpdateDependencyControls
        End Sub

        Private Sub PopulateDependencyCombo()
            Try
                _dependsOnCombo.Items.Clear()
                _dependsOnCombo.Items.Add("(None)")

                ' Get all tasks from scheduler
                Dim scheduler = ApplicationContainer.GetService(Of CustomScheduler)()
                For Each task In scheduler.GetAllTasks()
                    _dependsOnCombo.Items.Add(task.Name)
                Next

                _dependsOnCombo.SelectedIndex = 0

            Catch ex As Exception
                _logger.LogError($"Error populating dependency combo: {ex.Message}")
            End Try
        End Sub

        Private Sub UpdateDependencyControls(sender As Object, e As EventArgs)
            _timeoutNumeric.Enabled = _waitForCompletionCheck.Checked
            _retryCountNumeric.Enabled = _waitForCompletionCheck.Checked
            _retryDelayNumeric.Enabled = _waitForCompletionCheck.Checked
        End Sub

        Private Sub TypeComboBox_SelectedIndexChanged(sender As Object, e As EventArgs)
            _browseButton.Enabled = True
        End Sub

        Private Sub BrowseButton_Click(sender As Object, e As EventArgs)
            Try
                Using openFileDialog As New OpenFileDialog()
                    Select Case _typeComboBox.SelectedItem.ToString()
                        Case "PowerShell"
                            openFileDialog.Filter = "PowerShell Scripts (*.ps1)|*.ps1|All files (*.*)|*.*"
                        Case "SQL"
                            openFileDialog.Filter = "SQL Files (*.sql)|*.sql|All files (*.*)|*.*"
                        Case "Executable"
                            openFileDialog.Filter = "Executable Files (*.exe)|*.exe|All files (*.*)|*.*"
                        Case "Batch"
                            openFileDialog.Filter = "Batch Files (*.bat;*.cmd)|*.bat;*.cmd|All files (*.*)|*.*"
                        Case Else
                            openFileDialog.Filter = "All files (*.*)|*.*"
                    End Select

                    If openFileDialog.ShowDialog() = DialogResult.OK Then
                        _targetTextBox.Text = openFileDialog.FileName
                    End If
                End Using
            Catch ex As Exception
                _logger.LogError($"Error in file browse dialog: {ex.Message}")
                MessageBox.Show($"Error selecting file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub OkButton_Click(sender As Object, e As EventArgs)
            Try
                If Not ValidateInput() Then
                    DialogResult = DialogResult.None
                    Return
                End If

                ' Create or update action
                If Action Is Nothing Then
                    Action = New TaskAction()
                End If

                ' Set basic properties
                Action.Name = _nameTextBox.Text.Trim()
                Action.Type = CType([Enum].Parse(GetType(TaskType), _typeComboBox.SelectedItem.ToString()), TaskType)
                Action.Target = _targetTextBox.Text.Trim()
                Action.Parameters = _parametersTextBox.Text.Trim()
                Action.RequiresElevation = _requiresElevationCheckBox.Checked

                ' Set dependency properties
                Action.DependsOn = If(_dependsOnCombo.SelectedIndex > 0, _dependsOnCombo.SelectedItem.ToString(), Nothing)
                Action.WaitForCompletion = _waitForCompletionCheck.Checked
                Action.TimeoutMinutes = Convert.ToInt32(_timeoutNumeric.Value)
                Action.RetryCount = Convert.ToInt32(_retryCountNumeric.Value)
                Action.RetryDelayMinutes = Convert.ToInt32(_retryDelayNumeric.Value)
                Action.ContinueOnError = _continueOnErrorCheck.Checked

            Catch ex As Exception
                _logger.LogError($"Error saving action: {ex.Message}")
                MessageBox.Show($"Error saving action: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                DialogResult = DialogResult.None
            End Try
        End Sub

        Private Function ValidateInput() As Boolean
            If String.IsNullOrWhiteSpace(_nameTextBox.Text) Then
                MessageBox.Show("Please enter a name for the action.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If

            If String.IsNullOrWhiteSpace(_targetTextBox.Text) Then
                MessageBox.Show("Please specify a target.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If

            Return True
        End Function


    End Class
End Namespace