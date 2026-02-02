Imports System.Drawing
Imports LiteTask.LiteTask.ScheduledTask

Namespace LiteTask
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Public Class TaskForm
        Friend WithEvents buttonPanel As FlowLayoutPanel

        'Required by the Windows Form Designer
        Private components As System.ComponentModel.IContainer

        'Form overrides dispose to clean up the component list.
        <System.Diagnostics.DebuggerNonUserCode()>
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            Try
                If disposing AndAlso components IsNot Nothing Then
                    components.Dispose()
                End If
            Finally
                MyBase.Dispose(disposing)
            End Try
        End Sub


        ' Form Controls - These need to be Private for the designer but accessible through Protected Friend properties
        Private WithEvents _tabControl As TabControl
        Private WithEvents _generalTab As TabPage
        Private WithEvents _scheduleTab As TabPage
        Private WithEvents _actionsTab As TabPage

        ' General tab controls
        Private _nameLabel As Label
        Private _descriptionLabel As Label
        Private _credentialLabel As Label
        Private WithEvents _nameTextBox As TextBox
        Private WithEvents _descriptionTextBox As TextBox
        Private WithEvents _credentialComboBox As ComboBox
        Private WithEvents _requiresElevationCheckBox As CheckBox

        ' Schedule tab controls
        Private WithEvents _oneTimeRadio As RadioButton
        Private WithEvents _recurringRadio As RadioButton
        Private WithEvents _schedulePicker As DateTimePicker
        Private WithEvents _timePicker As DateTimePicker
        Private WithEvents _recurrenceTypeCombo As ComboBox
        Private WithEvents _intervalPanel As Panel
        Private WithEvents _intervalTextBox As TextBox
        Private WithEvents _intervalLabel As Label
        Private WithEvents _intervalUnitLabel As Label
        Private WithEvents _dailyPanel As Panel
        Private WithEvents _dailyTimeList As ListBox
        Private WithEvents _dailyTimePicker As DateTimePicker
        Private WithEvents _addTimeButton As Button
        Private WithEvents _removeTimeButton As Button
        Private _dailyTimesList As ListBox
        Private _recurringCheckBox As CheckBox
        Private _scheduleDatePicker As DateTimePicker
        Private _scheduleTimePicker As DateTimePicker

        ' Actions tab controls
        Private WithEvents _actionsGrid As DataGridView
        Private WithEvents _actionsButtonPanel As FlowLayoutPanel
        Private WithEvents _addActionButton As Button
        Private WithEvents _editActionButton As Button
        Private WithEvents _deleteActionButton As Button
        Private WithEvents _moveUpButton As Button
        Private WithEvents _moveDownButton As Button

        ' Form buttons
        Private WithEvents _okButton As Button
        Private WithEvents _cancelButton As Button


        'Required by the Windows Form Designer
        <System.Diagnostics.DebuggerStepThrough()>
        Private Sub InitializeComponent()
            'Try
            '_logger?.LogInfo("Starting TaskForm InitializeComponent")
            Dim resources As ComponentResourceManager = New ComponentResourceManager(GetType(TaskForm))

                ' Initialize all controls first
                _tabControl = New TabControl()
                _generalTab = New TabPage()
                _scheduleTab = New TabPage()
                _actionsTab = New TabPage()

                ' General tab controls
                _nameTextBox = New TextBox()
                _descriptionTextBox = New TextBox()
                _credentialComboBox = New ComboBox()
                _requiresElevationCheckBox = New CheckBox()

                ' Schedule tab controls
                _scheduleDatePicker = New DateTimePicker()
                _scheduleTimePicker = New DateTimePicker()
                _recurringCheckBox = New CheckBox()
                _recurrenceTypeCombo = New ComboBox()
                _intervalPanel = New Panel()
                _intervalTextBox = New TextBox()
                _intervalLabel = New Label()
                _intervalUnitLabel = New Label()
                _dailyPanel = New Panel()
                _dailyTimeList = New ListBox()
                _dailyTimePicker = New DateTimePicker()
                _addTimeButton = New Button()
                _removeTimeButton = New Button()

                ' Actions tab controls
                _actionsGrid = New DataGridView()
                _actionsButtonPanel = New FlowLayoutPanel()
                _addActionButton = New Button()
                _editActionButton = New Button()
                _deleteActionButton = New Button()
                _moveUpButton = New Button()
                _moveDownButton = New Button()

                ' Form buttons
                _okButton = New Button()
                _cancelButton = New Button()
                buttonPanel = New FlowLayoutPanel()

                ' Configure TabControl
                _tabControl.SuspendLayout()
                _generalTab.SuspendLayout()
                _scheduleTab.SuspendLayout()
                _actionsTab.SuspendLayout()
                CType(_actionsGrid, ISupportInitialize).BeginInit()
                buttonPanel.SuspendLayout()
                SuspendLayout()

                ' Configure TabControl
                _tabControl.Controls.Add(_generalTab)
                _tabControl.Controls.Add(_scheduleTab)
                _tabControl.Controls.Add(_actionsTab)
                _tabControl.Dock = DockStyle.Fill
                _tabControl.Location = New Point(0, 0)
                _tabControl.Name = "_tabControl"
                _tabControl.SelectedIndex = 0
                _tabControl.Size = New Size(584, 401)
                _tabControl.TabIndex = 0

                ' Configure General Tab
                _generalTab.Location = New Point(4, 24)
                _generalTab.Name = "_generalTab"
                _generalTab.Padding = New Padding(3)
                _generalTab.Size = New Size(576, 373)
                _generalTab.TabIndex = 0
                _generalTab.Text = "General"
                _generalTab.UseVisualStyleBackColor = True

                ' Configure Schedule Tab
                _scheduleTab.Location = New Point(4, 24)
                _scheduleTab.Name = "_scheduleTab"
                _scheduleTab.Padding = New Padding(3)
                _scheduleTab.Size = New Size(576, 373)
                _scheduleTab.TabIndex = 1
                _scheduleTab.Text = "Schedule"
                _scheduleTab.UseVisualStyleBackColor = True

                ' Configure Actions Tab
                _actionsTab.Location = New Point(4, 24)
                _actionsTab.Name = "_actionsTab"
                _actionsTab.Padding = New Padding(3)
                _actionsTab.Size = New Size(576, 373)
                _actionsTab.TabIndex = 2
                _actionsTab.Text = "Actions"
                _actionsTab.UseVisualStyleBackColor = True

                ' Configure General Tab Controls
                _nameTextBox.Location = New Point(120, 20)
                _nameTextBox.Size = New Size(400, 23)
                _nameTextBox.Name = "_nameTextBox"

                _descriptionTextBox.Location = New Point(120, 50)
                _descriptionTextBox.Size = New Size(400, 60)
                _descriptionTextBox.Multiline = True
                _descriptionTextBox.Name = "_descriptionTextBox"

                _credentialComboBox.Location = New Point(120, 120)
                _credentialComboBox.Size = New Size(400, 23)
                _credentialComboBox.DropDownStyle = ComboBoxStyle.DropDownList
                _credentialComboBox.Name = "_credentialComboBox"

                _requiresElevationCheckBox.Location = New Point(120, 150)
                _requiresElevationCheckBox.Size = New Size(400, 23)
                _requiresElevationCheckBox.Text = "Requires Elevation"
                _requiresElevationCheckBox.Name = "_requiresElevationCheckBox"

                ' Configure Schedule Tab Controls
                _scheduleDatePicker.Location = New Point(120, 20)
                _scheduleDatePicker.Size = New Size(200, 23)
                _scheduleDatePicker.Format = DateTimePickerFormat.Short
                _scheduleDatePicker.Name = "_scheduleDatePicker"

                _scheduleTimePicker.Location = New Point(330, 20)
                _scheduleTimePicker.Size = New Size(100, 23)
                _scheduleTimePicker.Format = DateTimePickerFormat.Time
                _scheduleTimePicker.ShowUpDown = True
                _scheduleTimePicker.Name = "_scheduleTimePicker"

                _recurringCheckBox.Location = New Point(120, 50)
                _recurringCheckBox.Size = New Size(200, 23)
                _recurringCheckBox.Text = "Recurring"
                _recurringCheckBox.Name = "_recurringCheckBox"

                _recurrenceTypeCombo.Location = New Point(120, 80)
                _recurrenceTypeCombo.Size = New Size(200, 23)
                _recurrenceTypeCombo.DropDownStyle = ComboBoxStyle.DropDownList
                _recurrenceTypeCombo.Enabled = False
                _recurrenceTypeCombo.Name = "_recurrenceTypeCombo"
                _recurrenceTypeCombo.Items.AddRange([Enum].GetNames(GetType(RecurrenceType)))

                ' Configure Interval Panel
                _intervalPanel.Location = New Point(120, 110)
                _intervalPanel.Size = New Size(400, 60)
                _intervalPanel.Visible = False
                _intervalPanel.Name = "_intervalPanel"

                _intervalLabel.Location = New Point(0, 3)
                _intervalLabel.Size = New Size(55, 23)
                _intervalLabel.Text = "Interval:"
                _intervalLabel.Name = "_intervalLabel"

                _intervalTextBox.Location = New Point(60, 0)
                _intervalTextBox.Size = New Size(80, 23)
                _intervalTextBox.Name = "_intervalTextBox"

                _intervalUnitLabel.Location = New Point(145, 3)
                _intervalUnitLabel.Size = New Size(60, 23)
                _intervalUnitLabel.Text = "minutes"
                _intervalUnitLabel.Name = "_intervalUnitLabel"

                _intervalPanel.Controls.Add(_intervalLabel)
                _intervalPanel.Controls.Add(_intervalTextBox)
                _intervalPanel.Controls.Add(_intervalUnitLabel)

                ' Configure Daily Panel
                _dailyPanel.Location = New Point(120, 110)
                _dailyPanel.Size = New Size(400, 200)
                _dailyPanel.Visible = False
                _dailyPanel.Name = "_dailyPanel"

                _dailyTimeList.Location = New Point(0, 0)
                _dailyTimeList.Size = New Size(200, 150)
                _dailyTimeList.Name = "_dailyTimeList"

                _dailyTimePicker.Location = New Point(210, 0)
                _dailyTimePicker.Size = New Size(100, 23)
                _dailyTimePicker.Format = DateTimePickerFormat.Time
                _dailyTimePicker.ShowUpDown = True
                _dailyTimePicker.Name = "_dailyTimePicker"

                _addTimeButton.Location = New Point(210, 30)
                _addTimeButton.Size = New Size(75, 23)
                _addTimeButton.Text = "Add"
                _addTimeButton.Name = "_addTimeButton"

                _removeTimeButton.Location = New Point(210, 60)
                _removeTimeButton.Size = New Size(75, 23)
                _removeTimeButton.Text = "Remove"
                _removeTimeButton.Name = "_removeTimeButton"

                _dailyPanel.Controls.Add(_dailyTimeList)
                _dailyPanel.Controls.Add(_dailyTimePicker)
                _dailyPanel.Controls.Add(_addTimeButton)
                _dailyPanel.Controls.Add(_removeTimeButton)

                ' Add Schedule Controls to Schedule Tab
                _scheduleTab.Controls.Add(_scheduleDatePicker)
                _scheduleTab.Controls.Add(_scheduleTimePicker)
                _scheduleTab.Controls.Add(_recurringCheckBox)
                _scheduleTab.Controls.Add(_recurrenceTypeCombo)
                _scheduleTab.Controls.Add(_intervalPanel)
                _scheduleTab.Controls.Add(_dailyPanel)

                ' Configure Actions Grid
                _actionsGrid = New DataGridView() With {
            .Dock = DockStyle.Fill,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .MultiSelect = False,
            .Name = "_actionsGrid"
        }

                _actionsGrid.Columns.AddRange(New DataGridViewColumn() {
            New DataGridViewTextBoxColumn() With {
                .HeaderText = "Order",
                .Name = "OrderColumn",
                .Width = 60,
                .ReadOnly = True
            },
            New DataGridViewTextBoxColumn() With {
                .HeaderText = "Type",
                .Name = "TypeColumn",
                .Width = 100,
                .ReadOnly = True
            },
            New DataGridViewTextBoxColumn() With {
                .HeaderText = "Target",
                .Name = "TargetColumn",
                .Width = 250,
                .ReadOnly = True
            },
            New DataGridViewTextBoxColumn() With {
                .HeaderText = "Parameters",
                .Name = "ParametersColumn",
                .Width = 200,
                .ReadOnly = True
            },
            New DataGridViewCheckBoxColumn() With {
                .HeaderText = "Requires Elevation",
                .Name = "ElevationColumn",
                .Width = 80,
                .ReadOnly = True
            }
        })

                ' Configure Actions Button Panel
                _actionsButtonPanel = New FlowLayoutPanel() With {
            .Dock = DockStyle.Bottom,
            .Height = 40,
            .FlowDirection = FlowDirection.RightToLeft,
            .Padding = New Padding(5),
            .Name = "_actionsButtonPanel"
        }

                _addActionButton = New Button() With {
            .Text = "Add",
            .Name = "_addActionButton"
        }
                _editActionButton = New Button() With {
            .Text = "Edit",
            .Name = "_editActionButton"
        }
                _deleteActionButton = New Button() With {
            .Text = "Delete",
            .Name = "_deleteActionButton"
        }
                _moveUpButton = New Button() With {
            .Text = "Move Up",
            .Name = "_moveUpButton"
        }
                _moveDownButton = New Button() With {
            .Text = "Move Down",
            .Name = "_moveDownButton"
        }

                _actionsButtonPanel.Controls.AddRange({
            _moveDownButton,
            _moveUpButton,
            _deleteActionButton,
            _editActionButton,
            _addActionButton
        })

                ' Add Actions Controls to Actions Tab
                _actionsTab.Controls.Add(_actionsGrid)
                _actionsTab.Controls.Add(_actionsButtonPanel)

                ' Configure Form Buttons
                buttonPanel = New FlowLayoutPanel() With {
            .Dock = DockStyle.Bottom,
            .Height = 40,
            .FlowDirection = FlowDirection.RightToLeft,
            .Padding = New Padding(5)
        }

                _okButton = New Button() With {
            .Text = "OK",
            .DialogResult = DialogResult.OK,
            .Size = New Size(75, 23)
        }

                _cancelButton = New Button() With {
            .Text = "Cancel",
            .DialogResult = DialogResult.Cancel,
            .Size = New Size(75, 23)
        }

                buttonPanel.Controls.AddRange({_cancelButton, _okButton})

                ' Configure Form
                AutoScaleDimensions = New SizeF(7.0F, 15.0F)
                AutoScaleMode = AutoScaleMode.Font
                ClientSize = New Size(584, 441)
                Controls.Add(_tabControl)
                Controls.Add(buttonPanel)
                FormBorderStyle = FormBorderStyle.FixedDialog
                MaximizeBox = False
                MinimizeBox = False
                Name = "TaskForm"
            StartPosition = FormStartPosition.CenterParent

            ' Create labels with names
            _nameLabel = New Label With {
            .Text = TranslationManager.Instance.GetTranslation("_nameLabel.Text"),
            .Location = New Point(20, 23),
            .AutoSize = True,
            .Name = "_nameLabel"
             }

            _descriptionLabel = New Label With {
            .Text = TranslationManager.Instance.GetTranslation("_descriptionLabel.Text"),
            .Location = New Point(20, 53),
            .AutoSize = True,
            .Name = "_descriptionLabel"
            }

            _credentialLabel = New Label With {
            .Text = TranslationManager.Instance.GetTranslation("_credentialLabel.Text"),
            .Location = New Point(20, 123),
            .AutoSize = True,
            .Name = "_credentialLabel"
            }

            ' Add labels to General Tab
            _generalTab.Controls.Add(_nameLabel)
            _generalTab.Controls.Add(_descriptionLabel)
            _generalTab.Controls.Add(_credentialLabel)

            ' Add controls to General Tab
            _generalTab.Controls.Add(_nameTextBox)
            _generalTab.Controls.Add(_descriptionTextBox)
            _generalTab.Controls.Add(_credentialComboBox)
            _generalTab.Controls.Add(_requiresElevationCheckBox)

            ' Add labels to Schedule Tab
            _scheduleTab.Controls.Add(New Label With {
            .Text = "Date:",
            .Location = New Point(20, 23),
            .AutoSize = True
        })
            _scheduleTab.Controls.Add(New Label With {
            .Text = "Time:",
            .Location = New Point(280, 23),
            .AutoSize = True
        })

                _tabControl.ResumeLayout(False)
                _generalTab.ResumeLayout(False)
                _scheduleTab.ResumeLayout(False)
                _actionsTab.ResumeLayout(False)
                CType(_actionsGrid, ISupportInitialize).EndInit()
                buttonPanel.ResumeLayout(False)
            ResumeLayout(False)
        End Sub

    End Class
End Namespace