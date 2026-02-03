Imports LiteTask.LiteTask.ScheduledTask

Namespace LiteTask
    Public Class TaskForm
        Inherits Form

        Private ReadOnly _credentialManager As CredentialManager
        Private ReadOnly _customScheduler As CustomScheduler
        Private ReadOnly _logger As Logger
        Private _isEditMode As Boolean
        Private _task As ScheduledTask



        Public Sub New(credentialManager As CredentialManager,
                  customScheduler As CustomScheduler,
                  logger As Logger,
                  Optional existingTask As ScheduledTask = Nothing)

            ' This call is required by the designer and must be first
            InitializeComponent()

            Try
                ' Store dependencies
                _credentialManager = credentialManager
                _customScheduler = customScheduler
                _logger = logger

                ' Determine if we're in edit mode
                _isEditMode = (existingTask IsNot Nothing)
                Me.Text = If(_isEditMode, "Edit Task", "Create Task")

                ' Initialize task based on mode
                If _isEditMode Then
                    ' Clone the existing task to avoid modifying the original until save
                    _task = existingTask.Clone()
                Else
                    ' Create new task with default values
                    _task = New ScheduledTask() With {
                    .Actions = New List(Of TaskAction)(),
                    .DailyTimes = New List(Of TimeSpan)(),
                    .Parameters = New Hashtable(),
                    .Schedule = RecurrenceType.OneTime,
                    .StartTime = DateTime.Now,
                    .AccountType = "Current User"
                }
                End If

                ' Initialize form controls with data
                InitializeFormData()
                InitializeActionsGrid()
                Me.Translate()
                TranslateComponents()
                ' Add event handlers
                AddEventHandlers()

                _logger?.LogInfo("TaskForm initialized successfully")

            Catch ex As Exception
                _logger?.LogError($"Error initializing TaskForm: {ex.Message}")
                _logger?.LogError($"StackTrace: {ex.StackTrace}")
                Throw New InvalidOperationException("Failed to initialize TaskForm", ex)
            End Try
        End Sub

        Private Sub ActionsGrid_SelectionChanged(sender As Object, e As EventArgs) Handles _actionsGrid.SelectionChanged
            UpdateActionButtons()
        End Sub

        Private Sub AddActionButton_Click(sender As Object, e As EventArgs)
            Using dialog As New ActionDialog()
                If dialog.ShowDialog() = DialogResult.OK Then
                    Dim action = dialog.Action
                    action.Order = _task.Actions.Count + 1
                    _task.Actions.Add(action)
                    RefreshActionsList()
                End If
            End Using
        End Sub

        Private Sub AddEventHandlers()
            ' Add handlers for schedule controls
            AddHandler _recurringCheckBox.CheckedChanged, AddressOf RecurringCheckBox_CheckedChanged
            AddHandler _recurrenceTypeCombo.SelectedIndexChanged, AddressOf RecurrenceTypeCombo_SelectedIndexChanged

            ' Add handlers for daily time controls
            AddHandler _addTimeButton.Click, AddressOf AddTimeButton_Click
            AddHandler _removeTimeButton.Click, AddressOf RemoveTimeButton_Click

            ' Add handlers for action controls
            AddHandler _addActionButton.Click, AddressOf AddActionButton_Click
            AddHandler _editActionButton.Click, AddressOf EditActionButton_Click
            AddHandler _deleteActionButton.Click, AddressOf DeleteActionButton_Click
            AddHandler _moveUpButton.Click, AddressOf MoveUpButton_Click
            AddHandler _moveDownButton.Click, AddressOf MoveDownButton_Click

            ' Add handlers for form buttons
            AddHandler _okButton.Click, AddressOf OkButton_Click
        End Sub

        Private Sub AddTimeButton_Click(sender As Object, e As EventArgs)
            Try
                ' Format time as 24-hour
                Dim timeString = _dailyTimePicker.Value.ToString("HH:mm")

                ' Check if this time already exists
                If Not _dailyTimeList.Items.Cast(Of String)().Any(Function(t) t = timeString) Then
                    _dailyTimeList.Items.Add(timeString)

                    ' Sort times
                    Dim times = _dailyTimeList.Items.Cast(Of String)().
                OrderBy(Function(t) TimeSpan.Parse(t)).
                ToArray()
                    _dailyTimeList.Items.Clear()
                    _dailyTimeList.Items.AddRange(times)

                    _logger?.LogInfo($"Added daily time: {timeString} to task schedule")
                Else
                    _logger?.LogInfo($"Time {timeString} already exists in schedule")
                End If
            Catch ex As Exception
                _logger?.LogError($"Error adding time: {ex.Message}")
                MessageBox.Show($"Error adding time: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub DeleteActionButton_Click(sender As Object, e As EventArgs)
            If _actionsGrid.SelectedRows.Count = 0 Then Return

            Dim index = _actionsGrid.SelectedRows(0).Index
            _task.Actions.RemoveAt(index)
            ReorderActions()
            LoadActions()
        End Sub

        Private Sub EditActionButton_Click(sender As Object, e As EventArgs)
            If _actionsGrid.SelectedRows.Count = 0 Then Return

            Dim index = _actionsGrid.SelectedRows(0).Index
            Dim action = _task.Actions(index)

            Using dialog As New ActionDialog(action)
                If dialog.ShowDialog() = DialogResult.OK Then
                    _task.Actions(index) = dialog.Action
                    RefreshActionsList()
                End If
            End Using
        End Sub

        Private Sub InitializeActionsGrid()
            _actionsGrid.AutoGenerateColumns = False
            _actionsGrid.Columns.Clear()
            _actionsGrid.Columns.AddRange(New DataGridViewColumn() {
        New DataGridViewTextBoxColumn() With {
            .DataPropertyName = "Order",
            .HeaderText = "Order",
            .Width = 60
        },
        New DataGridViewTextBoxColumn() With {
            .DataPropertyName = "Name",
            .HeaderText = "Name",
            .Width = 150
        },
        New DataGridViewTextBoxColumn() With {
            .DataPropertyName = "Type",
            .HeaderText = "Type",
            .Width = 100
        },
        New DataGridViewTextBoxColumn() With {
            .DataPropertyName = "Target",
            .HeaderText = "Target",
            .Width = 200
        }
    })
            RefreshActionsList()
        End Sub

        Private Sub InitializeFormData()
            Try
                ' Initialize credentials combo
                _credentialComboBox.Items.Clear()
                _credentialComboBox.Items.Add("(None)")
                _credentialComboBox.Items.AddRange(_credentialManager.GetAllCredentialTargets().ToArray())
                _credentialComboBox.SelectedIndex = 0

                ' Initialize recurrence type combo with only Daily and Interval
                _recurrenceTypeCombo.Items.Clear()
                ' Store the actual enum values as Tag property
                Dim dailyItem As New ComboBoxItem(
            TranslationManager.Instance.GetTranslation("Daily"),
            RecurrenceType.Daily)
                Dim intervalItem As New ComboBoxItem(
            TranslationManager.Instance.GetTranslation("Interval"),
            RecurrenceType.Interval)

                Dim monthlyItem As New ComboBoxItem(
            TranslationManager.Instance.GetTranslation("Monthly"),
            RecurrenceType.Monthly)

                _recurrenceTypeCombo.Items.Add(dailyItem)
                _recurrenceTypeCombo.Items.Add(intervalItem)
                _recurrenceTypeCombo.Items.Add(monthlyItem)
                _recurrenceTypeCombo.SelectedIndex = 0

                ' Set default values for new tasks or load existing task data
                If Not _isEditMode Then
                    _scheduleDatePicker.Value = DateTime.Now.Date
                    _scheduleTimePicker.Value = DateTime.Now
                    _recurringCheckBox.Checked = False
                    _recurrenceTypeCombo.Enabled = False
                Else
                    ' Load existing task data
                    _nameTextBox.Text = _task.Name
                    _descriptionTextBox.Text = _task.Description
                    _scheduleDatePicker.Value = _task.StartTime.Date
                    _scheduleTimePicker.Value = _task.StartTime

                    ' Set recurrence settings
                    _recurringCheckBox.Checked = (_task.Schedule <> RecurrenceType.OneTime)
                    _recurrenceTypeCombo.Enabled = _recurringCheckBox.Checked

                    If _recurringCheckBox.Checked Then
                        ' Find and select the matching schedule type
                        For i As Integer = 0 To _recurrenceTypeCombo.Items.Count - 1
                            Dim item = DirectCast(_recurrenceTypeCombo.Items(i), ComboBoxItem)
                            If item.Value = _task.Schedule Then
                                _recurrenceTypeCombo.SelectedIndex = i
                                Exit For
                            End If
                        Next

                        Select Case _task.Schedule
                            Case RecurrenceType.Interval
                                _intervalTextBox.Text = _task.Interval.TotalMinutes.ToString()

                            Case RecurrenceType.Daily
                                _dailyTimeList.Items.Clear()
                                For Each time In _task.DailyTimes
                                    Try
                                        _dailyTimeList.Items.Add(time.ToString("hh\:mm"))
                                    Catch ex As Exception
                                        _logger?.LogWarning($"Invalid time format in daily times: {time}")
                                    End Try
                                Next

                            Case RecurrenceType.Monthly
                                _monthlyDayNumeric.Value = If(_task.MonthlyDay >= 1 AndAlso _task.MonthlyDay <= 31, _task.MonthlyDay, 1)
                                _monthlyTimePicker.Value = DateTime.Today.Add(_task.MonthlyTime)
                        End Select
                    End If

                    ' Load credential settings
                    If Not String.IsNullOrEmpty(_task.CredentialTarget) Then
                        Dim credIndex = _credentialComboBox.Items.IndexOf(_task.CredentialTarget)
                        If credIndex >= 0 Then
                            _credentialComboBox.SelectedIndex = credIndex
                        End If
                    End If

                    _requiresElevationCheckBox.Checked = _task.RequiresElevation
                End If

                UpdateRecurrenceControls()
                _logger?.LogInfo($"Task form data initialized successfully. Edit mode: {_isEditMode}")

            Catch ex As Exception
                _logger?.LogError($"Error in InitializeFormData: {ex.Message}")
                _logger?.LogError($"StackTrace: {ex.StackTrace}")
                Throw
            End Try
        End Sub

        Private Sub LoadActions()
            _actionsGrid.Rows.Clear()
            For Each action In _task.Actions.OrderBy(Function(a) a.Order)
                _actionsGrid.Rows.Add(
                action.Order,
                action.Type.ToString(),
                action.Target,
                action.Parameters,
                action.RequiresElevation
            )
            Next
            UpdateActionButtons()
        End Sub

        Private Sub MoveDownButton_Click(sender As Object, e As EventArgs)
            If _actionsGrid.SelectedRows.Count = 0 OrElse _actionsGrid.SelectedRows(0).Index = _task.Actions.Count - 1 Then Return

            Dim index = _actionsGrid.SelectedRows(0).Index
            Dim action = _task.Actions(index)
            _task.Actions.RemoveAt(index)
            _task.Actions.Insert(index + 1, action)
            ReorderActions()
            LoadActions()
            _actionsGrid.Rows(index + 1).Selected = True
        End Sub

        Private Sub MoveUpButton_Click(sender As Object, e As EventArgs)
            If _actionsGrid.SelectedRows.Count = 0 OrElse _actionsGrid.SelectedRows(0).Index = 0 Then Return

            Dim index = _actionsGrid.SelectedRows(0).Index
            Dim action = _task.Actions(index)
            _task.Actions.RemoveAt(index)
            _task.Actions.Insert(index - 1, action)
            ReorderActions()
            LoadActions()
            _actionsGrid.Rows(index - 1).Selected = True
        End Sub

        Private Sub OkButton_Click(sender As Object, e As EventArgs)
            Try
                If Not ValidateInput() Then
                    DialogResult = DialogResult.None
                    Return
                End If

                UpdateTaskFromForm()

                If _isEditMode Then
                    _customScheduler.UpdateTask(_task)
                Else
                    _customScheduler.AddTask(_task)
                End If

                DialogResult = DialogResult.OK
                Close()

            Catch ex As Exception
                _logger?.LogError($"Error saving task: {ex.Message}")
                MessageBox.Show($"Error saving task: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                DialogResult = DialogResult.None
            End Try
        End Sub

        Private Sub RecurrenceTypeCombo_SelectedIndexChanged(sender As Object, e As EventArgs)
            UpdateRecurrenceControls()
        End Sub

        Private Sub RecurringCheckBox_CheckedChanged(sender As Object, e As EventArgs)
            _recurrenceTypeCombo.Enabled = _recurringCheckBox.Checked
            UpdateRecurrenceControls()
        End Sub

        Private Sub RefreshActionsList()
            Dim bindingSource = New BindingSource()
            bindingSource.DataSource = _task.Actions.OrderBy(Function(a) a.Order).ToList()
            _actionsGrid.DataSource = bindingSource
        End Sub

        Private Sub RemoveTimeButton_Click(sender As Object, e As EventArgs)
            If _dailyTimeList.SelectedIndex <> -1 Then
                Dim removedTime = _dailyTimeList.SelectedItem.ToString()
                _dailyTimeList.Items.RemoveAt(_dailyTimeList.SelectedIndex)
                _logger?.LogInfo($"Removed time {removedTime} from schedule")
            End If
        End Sub

        Private Sub ReorderActions()
            Dim orderedActions = _task.Actions.OrderBy(Function(a) a.Order).ToList()
            For i As Integer = 0 To orderedActions.Count - 1
                orderedActions(i).Order = i + 1
            Next
        End Sub

        Private Sub TranslateComponents()
            _generalTab.Text = TranslationManager.Instance.GetTranslation("_generalTab.Text")
            _nameLabel.Text = TranslationManager.Instance.GetTranslation("_nameLabel.Text")
            _descriptionLabel.Text = TranslationManager.Instance.GetTranslation("_descriptionLabel.Text")
            _credentialLabel.Text = TranslationManager.Instance.GetTranslation("_credentialLabel.Text")
            _cancelButton.Text = TranslationManager.Instance.GetTranslation("_cancelButton.Text")
        End Sub

        Private Sub UpdateActionButtons()
            Dim hasSelection = _actionsGrid.SelectedRows.Count > 0
            Dim selectedIndex = If(hasSelection, _actionsGrid.SelectedRows(0).Index, -1)

            _editActionButton.Enabled = hasSelection
            _deleteActionButton.Enabled = hasSelection
            _moveUpButton.Enabled = hasSelection AndAlso selectedIndex > 0
            _moveDownButton.Enabled = hasSelection AndAlso selectedIndex < _actionsGrid.Rows.Count - 1
        End Sub

        Private Sub UpdateRecurrenceControls()
            Try
                _intervalPanel.Visible = False
                _dailyPanel.Visible = False
                _monthlyPanel.Visible = False

                If _recurringCheckBox.Checked AndAlso _recurrenceTypeCombo.SelectedItem IsNot Nothing Then
                    Dim selectedItem = DirectCast(_recurrenceTypeCombo.SelectedItem, ComboBoxItem)
                    Select Case selectedItem.Value
                        Case RecurrenceType.Interval
                            _intervalPanel.Visible = True
                        Case RecurrenceType.Daily
                            _dailyPanel.Visible = True
                            _logger?.LogInfo("Daily time controls made visible")
                        Case RecurrenceType.Monthly
                            _monthlyPanel.Visible = True
                            _logger?.LogInfo("Monthly controls made visible")
                    End Select
                End If
            Catch ex As Exception
                _logger?.LogError($"Error updating recurrence controls: {ex.Message}")
                _logger?.LogError($"StackTrace: {ex.StackTrace}")
            End Try
        End Sub

        Private Sub UpdateTaskFromForm()
            Try
                _task.Name = _nameTextBox.Text.Trim()
                _task.Description = _descriptionTextBox.Text.Trim()

                ' Update schedule
                Dim scheduleDate = _scheduleDatePicker.Value.Date
                Dim scheduleTime = _scheduleTimePicker.Value.TimeOfDay
                _task.StartTime = scheduleDate.Add(scheduleTime)

                ' Update recurrence
                If _recurringCheckBox.Checked AndAlso _recurrenceTypeCombo.SelectedItem IsNot Nothing Then
                    Dim selectedItem = DirectCast(_recurrenceTypeCombo.SelectedItem, ComboBoxItem)
                    _task.Schedule = selectedItem.Value

                    Select Case _task.Schedule
                        Case RecurrenceType.Interval
                            _task.Interval = TimeSpan.FromMinutes(Double.Parse(_intervalTextBox.Text))
                            _task.DailyTimes.Clear()

                        Case RecurrenceType.Daily
                            _task.Interval = TimeSpan.Zero
                            _task.DailyTimes.Clear()
                            For Each timeStr As String In _dailyTimeList.Items
                                Dim timeSpan As TimeSpan
                                If TimeSpan.TryParse(timeStr, timeSpan) Then
                                    _task.DailyTimes.Add(timeSpan)
                                End If
                            Next
                            _task.DailyTimes.Sort()

                        Case RecurrenceType.Monthly
                            _task.Interval = TimeSpan.Zero
                            _task.DailyTimes.Clear()
                            _task.MonthlyDay = CInt(_monthlyDayNumeric.Value)
                            _task.MonthlyTime = _monthlyTimePicker.Value.TimeOfDay
                    End Select
                Else
                    _task.Schedule = RecurrenceType.OneTime
                    _task.Interval = TimeSpan.Zero
                    _task.DailyTimes.Clear()
                End If

                _task.CredentialTarget = If(_credentialComboBox.SelectedIndex > 0,
                                  _credentialComboBox.SelectedItem.ToString(),
                                  "")
                _task.RequiresElevation = _requiresElevationCheckBox.Checked

                For i As Integer = 0 To _task.Actions.Count - 1
                    _task.Actions(i).RequiresElevation = _task.Actions(i).RequiresElevation
                Next

                _task.NextRunTime = _task.CalculateNextRunTime()

            Catch ex As Exception
                _logger?.LogError($"Error updating task from form: {ex.Message}")
                Throw
            End Try
        End Sub

        Private Function ValidateInput() As Boolean
            If String.IsNullOrWhiteSpace(_nameTextBox.Text) Then
                MessageBox.Show("Task name is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If

            If _recurringCheckBox.Checked AndAlso _recurrenceTypeCombo.SelectedItem IsNot Nothing Then
                Dim selectedItem = DirectCast(_recurrenceTypeCombo.SelectedItem, ComboBoxItem)
                Select Case selectedItem.Value
                    Case RecurrenceType.Interval
                        Dim interval As Double
                        If Not Double.TryParse(_intervalTextBox.Text, interval) OrElse interval <= 0 Then
                            MessageBox.Show("Please enter a valid interval greater than zero.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            Return False
                        End If

                    Case RecurrenceType.Daily
                        If _dailyTimeList.Items.Count = 0 Then
                            MessageBox.Show("Please add at least one daily time.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            Return False
                        End If

                        _logger?.LogInfo($"Current daily times during validation: {String.Join(", ", _dailyTimeList.Items.Cast(Of String)())}")

                    Case RecurrenceType.Monthly
                        If _monthlyDayNumeric.Value < 1 OrElse _monthlyDayNumeric.Value > 31 Then
                            MessageBox.Show("Please select a valid day of the month (1-31).", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            Return False
                        End If
                End Select
            End If

            Return True
        End Function

        Private Class ComboBoxItem
            Public Property DisplayText As String
            Public Property Value As RecurrenceType

            Public Sub New(text As String, value As RecurrenceType)
                DisplayText = text
                Me.Value = value
            End Sub

            Public Overrides Function ToString() As String
                Return DisplayText
            End Function
        End Class

    End Class
End Namespace