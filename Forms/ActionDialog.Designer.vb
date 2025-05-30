Imports System.Drawing
Imports LiteTask.LiteTask.ScheduledTask

Namespace LiteTask
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class ActionDialog
        Inherits System.Windows.Forms.Form

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

        'Required by the Windows Form Designer
        Private components As System.ComponentModel.IContainer

        'NOTE: The following procedure is required by the Windows Form Designer
        'It can be modified using the Windows Form Designer.  
        'Do not modify it using the code editor.
        <System.Diagnostics.DebuggerStepThrough()>
        Public Sub InitializeComponent()
            Dim resources As ComponentResourceManager = New ComponentResourceManager(GetType(ActionDialog))
            _nameLabel = New Label()
            _nameTextBox = New TextBox()
            _typeLabel = New Label()
            _typeComboBox = New ComboBox()
            _targetLabel = New Label()
            _targetTextBox = New TextBox()
            _browseButton = New Button()
            _parametersLabel = New Label()
            _parametersTextBox = New TextBox()
            _dependsOnLabel = New Label()
            _dependsOnCombo = New ComboBox()
            _waitForCompletionCheck = New CheckBox()
            _timeoutLabel = New Label()
            _timeoutNumeric = New NumericUpDown()
            _retryCountLabel = New Label()
            _retryCountNumeric = New NumericUpDown()
            _retryDelayLabel = New Label()
            _retryDelayNumeric = New NumericUpDown()
            _continueOnErrorCheck = New CheckBox()
            _okButton = New Button()
            _cancelButton = New Button()
            _requiresElevationCheckBox = New CheckBox()
            CType(_timeoutNumeric, ISupportInitialize).BeginInit()
            CType(_retryCountNumeric, ISupportInitialize).BeginInit()
            CType(_retryDelayNumeric, ISupportInitialize).BeginInit()
            SuspendLayout()
            ' 
            ' _nameLabel
            ' 
            _nameLabel.AutoSize = True
            _nameLabel.Location = New Point(14, 17)
            _nameLabel.Margin = New Padding(4, 0, 4, 0)
            _nameLabel.Name = "_nameLabel"
            _nameLabel.Size = New Size(42, 15)
            _nameLabel.TabIndex = 0
            _nameLabel.Text = "Name:"
            ' 
            ' _nameTextBox
            ' 
            _nameTextBox.Location = New Point(128, 14)
            _nameTextBox.Margin = New Padding(4, 3, 4, 3)
            _nameTextBox.Name = "_nameTextBox"
            _nameTextBox.Size = New Size(305, 23)
            _nameTextBox.TabIndex = 1
            ' 
            ' _typeLabel
            ' 
            _typeLabel.AutoSize = True
            _typeLabel.Location = New Point(14, 47)
            _typeLabel.Margin = New Padding(4, 0, 4, 0)
            _typeLabel.Name = "_typeLabel"
            _typeLabel.Size = New Size(34, 15)
            _typeLabel.TabIndex = 2
            _typeLabel.Text = "Type:"
            ' 
            ' _typeComboBox
            ' 
            _typeComboBox.DropDownStyle = ComboBoxStyle.DropDownList
            _typeComboBox.FormattingEnabled = True
            _typeComboBox.Location = New Point(128, 44)
            _typeComboBox.Margin = New Padding(4, 3, 4, 3)
            _typeComboBox.Name = "_typeComboBox"
            _typeComboBox.Size = New Size(140, 23)
            _typeComboBox.TabIndex = 3
            ' 
            ' _targetLabel
            ' 
            _targetLabel.AutoSize = True
            _targetLabel.Location = New Point(14, 78)
            _targetLabel.Margin = New Padding(4, 0, 4, 0)
            _targetLabel.Name = "_targetLabel"
            _targetLabel.Size = New Size(42, 15)
            _targetLabel.TabIndex = 4
            _targetLabel.Text = "Target:"
            ' 
            ' _targetTextBox
            ' 
            _targetTextBox.Location = New Point(128, 75)
            _targetTextBox.Margin = New Padding(4, 3, 4, 3)
            _targetTextBox.Name = "_targetTextBox"
            _targetTextBox.Size = New Size(305, 23)
            _targetTextBox.TabIndex = 5
            ' 
            ' _browseButton
            ' 
            _browseButton.Location = New Point(346, 73)
            _browseButton.Margin = New Padding(4, 3, 4, 3)
            _browseButton.Name = "_browseButton"
            _browseButton.Size = New Size(88, 27)
            _browseButton.TabIndex = 6
            _browseButton.Text = "Browse"
            _browseButton.UseVisualStyleBackColor = True
            ' 
            ' _parametersLabel
            ' 
            _parametersLabel.AutoSize = True
            _parametersLabel.Location = New Point(14, 108)
            _parametersLabel.Margin = New Padding(4, 0, 4, 0)
            _parametersLabel.Name = "_parametersLabel"
            _parametersLabel.Size = New Size(69, 15)
            _parametersLabel.TabIndex = 7
            _parametersLabel.Text = "Parameters:"
            ' 
            ' _parametersTextBox
            ' 
            _parametersTextBox.Location = New Point(128, 105)
            _parametersTextBox.Margin = New Padding(4, 3, 4, 3)
            _parametersTextBox.Name = "_parametersTextBox"
            _parametersTextBox.Size = New Size(305, 23)
            _parametersTextBox.TabIndex = 8
            ' 
            ' _dependsOnLabel
            ' 
            _dependsOnLabel.AutoSize = True
            _dependsOnLabel.Location = New Point(14, 138)
            _dependsOnLabel.Margin = New Padding(4, 0, 4, 0)
            _dependsOnLabel.Name = "_dependsOnLabel"
            _dependsOnLabel.Size = New Size(73, 15)
            _dependsOnLabel.TabIndex = 9
            _dependsOnLabel.Text = "Depends on:"
            ' 
            ' _dependsOnCombo
            ' 
            _dependsOnCombo.DropDownStyle = ComboBoxStyle.DropDownList
            _dependsOnCombo.FormattingEnabled = True
            _dependsOnCombo.Location = New Point(128, 135)
            _dependsOnCombo.Margin = New Padding(4, 3, 4, 3)
            _dependsOnCombo.Name = "_dependsOnCombo"
            _dependsOnCombo.Size = New Size(305, 23)
            _dependsOnCombo.TabIndex = 10
            ' 
            ' _waitForCompletionCheck
            ' 
            _waitForCompletionCheck.AutoSize = True
            _waitForCompletionCheck.Location = New Point(18, 166)
            _waitForCompletionCheck.Margin = New Padding(4, 3, 4, 3)
            _waitForCompletionCheck.Name = "_waitForCompletionCheck"
            _waitForCompletionCheck.Size = New Size(132, 19)
            _waitForCompletionCheck.TabIndex = 11
            _waitForCompletionCheck.Text = "Wait for completion"
            _waitForCompletionCheck.UseVisualStyleBackColor = True
            ' 
            ' _timeoutLabel
            ' 
            _timeoutLabel.AutoSize = True
            _timeoutLabel.Location = New Point(31, 193)
            _timeoutLabel.Margin = New Padding(4, 0, 4, 0)
            _timeoutLabel.Name = "_timeoutLabel"
            _timeoutLabel.Size = New Size(54, 15)
            _timeoutLabel.TabIndex = 12
            _timeoutLabel.Text = "Timeout:"
            ' 
            ' _timeoutNumeric
            ' 
            _timeoutNumeric.Location = New Point(128, 190)
            _timeoutNumeric.Margin = New Padding(4, 3, 4, 3)
            _timeoutNumeric.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
            _timeoutNumeric.Name = "_timeoutNumeric"
            _timeoutNumeric.Size = New Size(88, 23)
            _timeoutNumeric.TabIndex = 13
            ' 
            ' _retryCountLabel
            ' 
            _retryCountLabel.AutoSize = True
            _retryCountLabel.Location = New Point(31, 223)
            _retryCountLabel.Margin = New Padding(4, 0, 4, 0)
            _retryCountLabel.Name = "_retryCountLabel"
            _retryCountLabel.Size = New Size(71, 15)
            _retryCountLabel.TabIndex = 14
            _retryCountLabel.Text = "Retry count:"
            ' 
            ' _retryCountNumeric
            ' 
            _retryCountNumeric.Location = New Point(128, 220)
            _retryCountNumeric.Margin = New Padding(4, 3, 4, 3)
            _retryCountNumeric.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
            _retryCountNumeric.Name = "_retryCountNumeric"
            _retryCountNumeric.Size = New Size(88, 23)
            _retryCountNumeric.TabIndex = 15
            ' 
            ' _retryDelayLabel
            ' 
            _retryDelayLabel.AutoSize = True
            _retryDelayLabel.Location = New Point(31, 253)
            _retryDelayLabel.Margin = New Padding(4, 0, 4, 0)
            _retryDelayLabel.Name = "_retryDelayLabel"
            _retryDelayLabel.Size = New Size(68, 15)
            _retryDelayLabel.TabIndex = 16
            _retryDelayLabel.Text = "Retry delay:"
            ' 
            ' _retryDelayNumeric
            ' 
            _retryDelayNumeric.Location = New Point(128, 250)
            _retryDelayNumeric.Margin = New Padding(4, 3, 4, 3)
            _retryDelayNumeric.Maximum = New Decimal(New Integer() {60, 0, 0, 0})
            _retryDelayNumeric.Name = "_retryDelayNumeric"
            _retryDelayNumeric.Size = New Size(88, 23)
            _retryDelayNumeric.TabIndex = 17
            ' 
            ' _continueOnErrorCheck
            ' 
            _continueOnErrorCheck.AutoSize = True
            _continueOnErrorCheck.Location = New Point(18, 280)
            _continueOnErrorCheck.Margin = New Padding(4, 3, 4, 3)
            _continueOnErrorCheck.Name = "_continueOnErrorCheck"
            _continueOnErrorCheck.Size = New Size(120, 19)
            _continueOnErrorCheck.TabIndex = 18
            _continueOnErrorCheck.Text = "Continue on error"
            _continueOnErrorCheck.UseVisualStyleBackColor = True
            ' 
            ' _okButton
            ' 
            _okButton.DialogResult = DialogResult.OK
            _okButton.Location = New Point(252, 307)
            _okButton.Margin = New Padding(4, 3, 4, 3)
            _okButton.Name = "_okButton"
            _okButton.Size = New Size(88, 27)
            _okButton.TabIndex = 19
            _okButton.Text = "OK"
            _okButton.UseVisualStyleBackColor = True
            ' 
            ' _cancelButton
            ' 
            _cancelButton.DialogResult = DialogResult.Cancel
            _cancelButton.Location = New Point(346, 307)
            _cancelButton.Margin = New Padding(4, 3, 4, 3)
            _cancelButton.Name = "_cancelButton"
            _cancelButton.Size = New Size(88, 27)
            _cancelButton.TabIndex = 20
            _cancelButton.Text = "Cancel"
            _cancelButton.UseVisualStyleBackColor = True
            ' 
            ' _requiresElevationCheckBox
            ' 
            _requiresElevationCheckBox.AutoSize = True
            _requiresElevationCheckBox.Location = New Point(18, 307)
            _requiresElevationCheckBox.Margin = New Padding(4, 3, 4, 3)
            _requiresElevationCheckBox.Name = "_requiresElevationCheckBox"
            _requiresElevationCheckBox.Size = New Size(122, 19)
            _requiresElevationCheckBox.TabIndex = 21
            _requiresElevationCheckBox.Text = "Requires Elevation"
            _requiresElevationCheckBox.UseVisualStyleBackColor = True
            ' 
            ' ActionDialog
            ' 
            AutoScaleDimensions = New SizeF(7.0F, 15.0F)
            AutoScaleMode = AutoScaleMode.Font
            ClientSize = New Size(448, 347)
            Controls.Add(_requiresElevationCheckBox)
            Controls.Add(_cancelButton)
            Controls.Add(_okButton)
            Controls.Add(_continueOnErrorCheck)
            Controls.Add(_retryDelayNumeric)
            Controls.Add(_retryDelayLabel)
            Controls.Add(_retryCountNumeric)
            Controls.Add(_retryCountLabel)
            Controls.Add(_timeoutNumeric)
            Controls.Add(_timeoutLabel)
            Controls.Add(_waitForCompletionCheck)
            Controls.Add(_dependsOnCombo)
            Controls.Add(_dependsOnLabel)
            Controls.Add(_parametersTextBox)
            Controls.Add(_parametersLabel)
            Controls.Add(_browseButton)
            Controls.Add(_targetTextBox)
            Controls.Add(_targetLabel)
            Controls.Add(_typeComboBox)
            Controls.Add(_typeLabel)
            Controls.Add(_nameTextBox)
            Controls.Add(_nameLabel)
            Icon = CType(resources.GetObject("$this.Icon"), Icon)
            Margin = New Padding(4, 3, 4, 3)
            MaximizeBox = False
            MinimizeBox = False
            Name = "ActionDialog"
            StartPosition = FormStartPosition.CenterParent
            Text = "Action Properties"
            CType(_timeoutNumeric, ISupportInitialize).EndInit()
            CType(_retryCountNumeric, ISupportInitialize).EndInit()
            CType(_retryDelayNumeric, ISupportInitialize).EndInit()
            ResumeLayout(False)
            PerformLayout()

        End Sub


    End Class
End Namespace