Imports System.Drawing

Namespace LiteTask
    Partial Public Class OptionsForm
        Inherits Form

        Private components As System.ComponentModel.IContainer

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        ' Tab Controls
        Private WithEvents _tabControl As TabControl
        Private WithEvents _loggingTab As TabPage
        Private WithEvents _notificationsTab As TabPage
        Private WithEvents _sqlTab As TabPage

        ' Logging Tab Controls
        Private WithEvents _logFolderLabel As Label
        Private WithEvents _logLevelLabel As Label
        Private WithEvents _alertLevelLabel As Label
        Private WithEvents _maxLogSizeLabel As Label
        Private WithEvents _logRetentionLabel As Label
        Private WithEvents _logFolderTextBox As TextBox
        Private WithEvents _logLevelComboBox As ComboBox
        Private WithEvents _browseLogFolderButton As Button
        Private WithEvents _maxLogSizeNumeric As NumericUpDown
        Private WithEvents _logRetentionNumeric As NumericUpDown
        Private WithEvents _alertLevelComboBox As ComboBox

        ' Email Tab Controls
        Private WithEvents _emailNotificationsLabel As Label
        Private WithEvents _smtpServerLabel As Label
        Private WithEvents _smtpPortLabel As Label
        Private WithEvents _emailFromLabel As Label
        Private WithEvents _emailToLabel As Label
        Private WithEvents _enableNotificationsCheckBox As CheckBox
        Private WithEvents _smtpServerTextBox As TextBox
        Private WithEvents _smtpPortNumeric As NumericUpDown
        Private WithEvents _useSSLCheckBox As CheckBox
        Private WithEvents _emailFromTextBox As TextBox
        Private WithEvents _emailToTextBox As TextBox
        Private WithEvents _useCredentialsCheckBox As CheckBox
        Private WithEvents _credentialComboBox As ComboBox
        Private WithEvents _testEmailButton As Button

        ' SQL Tab Controls
        Private WithEvents _defaultServerTextBox As TextBox
        Private WithEvents _defaultDatabaseTextBox As TextBox
        Private WithEvents _maxBatchSizeLabel As Label
        Private WithEvents _defaultServerLabel As Label
        Private WithEvents _defaultDatabaseLabel As Label
        Private WithEvents _commandTimeoutLabel As Label
        Private WithEvents _sqlPanel As TableLayoutPanel
        Private WithEvents _commandTimeoutNumeric As NumericUpDown
        Private WithEvents _maxBatchSizeNumeric As NumericUpDown

        ' Button Panel Controls
        Private WithEvents _buttonPanel As Panel
        Private WithEvents _okButton As Button
        Private WithEvents _cancelButton As Button

        Private Sub InitializeComponent()
            Me.components = New System.ComponentModel.Container()

            ' Initialize tab controls
            Me._tabControl = New TabControl()
            Me._loggingTab = New TabPage()
            Me._notificationsTab = New TabPage()
            Me._sqlTab = New TabPage()
            Me._buttonPanel = New Panel()

            ' Initialize Logging Tab Controls
            Me._logFolderLabel = New Label()
            Me._logLevelLabel = New Label()
            Me._alertLevelLabel = New Label()
            Me._maxLogSizeLabel = New Label()
            Me._logRetentionLabel = New Label()
            Me._logFolderTextBox = New TextBox()
            Me._logLevelComboBox = New ComboBox()
            Me._browseLogFolderButton = New Button()
            Me._maxLogSizeNumeric = New NumericUpDown()
            Me._logRetentionNumeric = New NumericUpDown()
            Me._alertLevelComboBox = New ComboBox()

            ' Initialize Email Tab Controls
            Me._emailNotificationsLabel = New Label()
            Me._smtpServerLabel = New Label()
            Me._smtpPortLabel = New Label()
            Me._emailFromLabel = New Label()
            Me._emailToLabel = New Label()
            Me._enableNotificationsCheckBox = New CheckBox()
            Me._smtpServerTextBox = New TextBox()
            Me._smtpPortNumeric = New NumericUpDown()
            Me._useSSLCheckBox = New CheckBox()
            Me._emailFromTextBox = New TextBox()
            Me._emailToTextBox = New TextBox()
            Me._useCredentialsCheckBox = New CheckBox()
            Me._credentialComboBox = New ComboBox()
            Me._testEmailButton = New Button()

            ' Initialize SQL Tab Controls
            Me._defaultServerTextBox = New TextBox()
            Me._defaultDatabaseTextBox = New TextBox()
            Me._maxBatchSizeLabel = New Label()
            Me._defaultServerLabel = New Label()
            Me._defaultDatabaseLabel = New Label()
            Me._commandTimeoutLabel = New Label()
            Me._sqlPanel = New TableLayoutPanel()
            Me._commandTimeoutNumeric = New NumericUpDown()
            Me._maxBatchSizeNumeric = New NumericUpDown()

            ' Initialize Buttons
            Me._okButton = New Button()
            Me._cancelButton = New Button()

            ' Configure Form
            Me.SuspendLayout()
            Me.Text = "Options"
            Me.Size = New Size(600, 450)
            Me.MinimizeBox = False
            Me.MaximizeBox = False
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.StartPosition = FormStartPosition.CenterParent

            ' Configure TabControl
            Me._tabControl.Dock = DockStyle.Fill
            Me._tabControl.Location = New Point(0, 0)
            Me._tabControl.Size = New Size(584, 362)

            ' Configure Button Panel
            Me._buttonPanel.Dock = DockStyle.Bottom
            Me._buttonPanel.Height = 50
            Me._buttonPanel.Padding = New Padding(10)

            Me._okButton.Text = "OK"
            Me._okButton.DialogResult = DialogResult.OK
            Me._okButton.Size = New Size(75, 23)
            Me._okButton.Location = New Point(Me._buttonPanel.Width - 170, 13)

            Me._cancelButton.Text = "Cancel"
            Me._cancelButton.DialogResult = DialogResult.Cancel
            Me._cancelButton.Size = New Size(75, 23)
            Me._cancelButton.Location = New Point(Me._buttonPanel.Width - 85, 13)

            Me._buttonPanel.Controls.Add(Me._okButton)
            Me._buttonPanel.Controls.Add(Me._cancelButton)

            ' Add main controls to form
            Me.Controls.Add(Me._tabControl)
            Me.Controls.Add(Me._buttonPanel)

            Me.ResumeLayout(False)
        End Sub
    End Class
End Namespace