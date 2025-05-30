Imports System.Drawing

Namespace LiteTask

    Partial Public Class OptionsForm

        ' Service/Manager declarations
        Private ReadOnly _xmlManager As XMLManager
        Private ReadOnly _logger As Logger
        Private ReadOnly _credentialManager As CredentialManager
        Private ReadOnly _errorHandler As ErrorHandler
        Private ReadOnly _settingValidator As SettingsValidator

        Public Sub New(xmlManager As XMLManager)
            Try
                _xmlManager = xmlManager
                _logger = ApplicationContainer.GetService(Of Logger)()
                _credentialManager = ApplicationContainer.GetService(Of CredentialManager)()
                _errorHandler = New ErrorHandler(_logger, ApplicationContainer.GetService(Of EmailUtils)())
                _settingValidator = New SettingsValidator(_logger)

                ' This call is required by the designer
                InitializeComponent()

                ' Initialize tabs
                InitializeLoggingTab()
                InitializeEmailTab()
                InitializeSqlTab()

                ' Apply translations and load settings
                ApplyTranslations()
                AddEventHandlers()
                LoadSettings()
            Catch ex As Exception
                _logger?.LogError($"Error initializing OptionsForm: {ex.Message}")
                MessageBox.Show($"Error initializing settings form: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()
            End Try
        End Sub

        Private Sub AddEventHandlers()
            AddHandler _okButton.Click, AddressOf SaveSettings
            AddHandler _enableNotificationsCheckBox.CheckedChanged, AddressOf UpdateEmailControls
            AddHandler _useCredentialsCheckBox.CheckedChanged, AddressOf UpdateCredentialControls
            AddHandler _testEmailButton.Click, AddressOf TestEmailButton_Click
            AddHandler _browseLogFolderButton.Click, AddressOf BrowseLogFolder_Click
            AddHandler _cancelButton.Click, AddressOf CancelButton_Click
        End Sub

        Private Sub ApplyTranslations()
            Try
                ' Form title
                Text = TranslationManager.Instance.GetTranslation("OptionsForm")

                ' Tabs
                _loggingTab.Text = TranslationManager.Instance.GetTranslation("_loggingTab")
                _notificationsTab.Text = TranslationManager.Instance.GetTranslation("_notificationsTab")
                _sqlTab.Text = TranslationManager.Instance.GetTranslation("_sqlTab")

                ' Logging tab controls
                _logFolderLabel.Text = TranslationManager.Instance.GetTranslation("_logFolderLabel")
                _logLevelLabel.Text = TranslationManager.Instance.GetTranslation("_logLevelLabel")
                _alertLevelLabel.Text = TranslationManager.Instance.GetTranslation("_alertLevelLabel")
                _maxLogSizeLabel.Text = TranslationManager.Instance.GetTranslation("_maxLogSizeLabel")
                _logRetentionLabel.Text = TranslationManager.Instance.GetTranslation("_logRetentionLabel")
                _browseLogFolderButton.Text = TranslationManager.Instance.GetTranslation("_browseLogFolderButton")

                ' Email tab controls
                _emailNotificationsLabel.Text = TranslationManager.Instance.GetTranslation("_emailNotificationsLabel")
                _enableNotificationsCheckBox.Text = TranslationManager.Instance.GetTranslation("_enableNotificationsCheckBox")
                _smtpServerLabel.Text = TranslationManager.Instance.GetTranslation("_smtpServerLabel")
                _smtpPortLabel.Text = TranslationManager.Instance.GetTranslation("_smtpPortLabel")
                _emailFromLabel.Text = TranslationManager.Instance.GetTranslation("_emailFromLabel")
                _emailToLabel.Text = TranslationManager.Instance.GetTranslation("_emailToLabel")
                _useSSLCheckBox.Text = TranslationManager.Instance.GetTranslation("_useSSLCheckBox")
                _useCredentialsCheckBox.Text = TranslationManager.Instance.GetTranslation("_useCredentialsCheckBox")
                _testEmailButton.Text = TranslationManager.Instance.GetTranslation("_testEmailButton")

                ' SQL tab controls
                _defaultServerLabel.Text = TranslationManager.Instance.GetTranslation("_defaultServerLabel")
                _defaultDatabaseLabel.Text = TranslationManager.Instance.GetTranslation("_defaultDatabaseLabel")
                _commandTimeoutLabel.Text = TranslationManager.Instance.GetTranslation("_commandTimeoutLabel")
                _maxBatchSizeLabel.Text = TranslationManager.Instance.GetTranslation("_maxBatchSizeLabel")

                ' Buttons
                _okButton.Text = TranslationManager.Instance.GetTranslation("_okButton")
                _cancelButton.Text = TranslationManager.Instance.GetTranslation("_cancelButton")

                ' Add tooltips
                Dim toolTip As New ToolTip()
                toolTip.SetToolTip(_logFolderTextBox, TranslationManager.Instance.GetTranslation("Options.Tooltip.LogFolder"))
                toolTip.SetToolTip(_logLevelComboBox, TranslationManager.Instance.GetTranslation("Options.Tooltip.LogLevel"))
                toolTip.SetToolTip(_alertLevelComboBox, TranslationManager.Instance.GetTranslation("Options.Tooltip.AlertLevel"))
                toolTip.SetToolTip(_maxLogSizeNumeric, TranslationManager.Instance.GetTranslation("Options.Tooltip.MaxLogSize"))
                toolTip.SetToolTip(_logRetentionNumeric, TranslationManager.Instance.GetTranslation("Options.Tooltip.RetentionDays"))
                toolTip.SetToolTip(_smtpServerTextBox, TranslationManager.Instance.GetTranslation("Options.Tooltip.SmtpServer"))
                toolTip.SetToolTip(_smtpPortNumeric, TranslationManager.Instance.GetTranslation("Options.Tooltip.SmtpPort"))
                toolTip.SetToolTip(_emailFromTextBox, TranslationManager.Instance.GetTranslation("Options.Tooltip.EmailFrom"))
                toolTip.SetToolTip(_emailToTextBox, TranslationManager.Instance.GetTranslation("Options.Tooltip.EmailTo"))
                toolTip.SetToolTip(_defaultServerLabel, TranslationManager.Instance.GetTranslation("Options.Tooltip.Server"))
                toolTip.SetToolTip(_defaultDatabaseLabel, TranslationManager.Instance.GetTranslation("Options.Tooltip.Database"))
                toolTip.SetToolTip(_commandTimeoutLabel, TranslationManager.Instance.GetTranslation("Options.Tooltip.Timeout"))
                toolTip.SetToolTip(_maxBatchSizeLabel, TranslationManager.Instance.GetTranslation("Options.Tooltip.BatchSize"))

            Catch ex As Exception
                _logger?.LogError($"Error applying translations: {ex.Message}")
                Throw
            End Try
        End Sub

        Private Sub BrowseLogFolder_Click(sender As Object, e As EventArgs)
            Using dialog As New FolderBrowserDialog()
                dialog.Description = "Select Log Folder"
                dialog.ShowNewFolderButton = True

                If Not String.IsNullOrEmpty(_logFolderTextBox.Text) AndAlso Directory.Exists(_logFolderTextBox.Text) Then
                    dialog.SelectedPath = _logFolderTextBox.Text
                End If

                If dialog.ShowDialog() = DialogResult.OK Then
                    _logFolderTextBox.Text = dialog.SelectedPath
                End If
            End Using
        End Sub

        Private Sub CancelButton_Click(sender As Object, e As EventArgs)
            DialogResult = DialogResult.Cancel
            Close()
        End Sub

        Private Function GetCurrentLogSettings() As Dictionary(Of String, String)
            Return New Dictionary(Of String, String) From {
        {"LogFolder", _logFolderTextBox.Text},
        {"LogLevel", _logLevelComboBox.SelectedItem.ToString()},
        {"MaxLogSize", _maxLogSizeNumeric.Value.ToString()},
        {"LogRetentionDays", _logRetentionNumeric.Value.ToString()},
        {"AlertLevel", _alertLevelComboBox.SelectedItem.ToString()}
    }
        End Function

        Private Function GetSettingValue(settings As Dictionary(Of String, String), key As String, defaultValue As String) As String
            If settings IsNot Nothing AndAlso settings.ContainsKey(key) Then
                Return settings(key)
            End If
            Return defaultValue
        End Function

        Private Sub HandleUnhandledException(sender As Object, e As ThreadExceptionEventArgs)
            _errorHandler.HandleException(e.Exception, "OptionsForm", True)
        End Sub

        Private Sub InitializeEmailTab()
            Dim emailTab As New TabPage("Email")
            Dim layout As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .Padding = New Padding(10),
            .ColumnCount = 2,
            .RowCount = 9
        }

            ' Set column styles
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 30))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 70))

            ' Initialize labels
            _emailNotificationsLabel = New Label With {
                .Text = "Email Notifications",
                .Dock = DockStyle.Fill
            }

            _enableNotificationsCheckBox = New CheckBox With {
                .Text = "Enable Email Notifications",
                .Dock = DockStyle.Fill
            }

            _smtpServerLabel = New Label With {
                .Text = "SMTP Server:",
                .Dock = DockStyle.Fill
            }

            _smtpPortLabel = New Label With {
                .Text = "SMTP Port:",
                .Dock = DockStyle.Fill
            }

            ' Initialize controls
            _enableNotificationsCheckBox = New CheckBox With {
            .Text = "Enable Email Notifications",
            .Dock = DockStyle.Fill
        }

            _smtpServerTextBox = New TextBox With {
            .Dock = DockStyle.Fill
        }

            _smtpPortNumeric = New NumericUpDown With {
            .Minimum = 1,
            .Maximum = 65535,
            .Value = 25,
            .Dock = DockStyle.Fill
        }

            _useSSLCheckBox = New CheckBox With {
            .Text = "Use SSL",
            .Dock = DockStyle.Fill
        }

            _emailFromTextBox = New TextBox With {
            .Dock = DockStyle.Fill
        }

            _emailToTextBox = New TextBox With {
            .Dock = DockStyle.Fill
        }

            _useCredentialsCheckBox = New CheckBox With {
            .Text = "Use Credentials",
            .Dock = DockStyle.Fill
        }

            _credentialComboBox = New ComboBox With {
            .Dock = DockStyle.Fill,
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Enabled = False
        }

            _testEmailButton = New Button With {
            .Text = "Send Test Email",
            .Dock = DockStyle.Fill
        }

            ' Populate credential combo
            _credentialComboBox.Items.Add("(None)")
            For Each cred In _credentialManager.GetAllCredentialTargets()
                _credentialComboBox.Items.Add(cred)
            Next
            _credentialComboBox.SelectedIndex = 0

            ' Add controls to layout
            layout.Controls.Add(_enableNotificationsCheckBox, 0, 0)
            layout.SetColumnSpan(_enableNotificationsCheckBox, 2)
            layout.Controls.Add(_smtpServerLabel, 0, 1)
            layout.Controls.Add(_smtpServerTextBox, 1, 1)
            layout.Controls.Add(_smtpPortLabel, 0, 2)
            layout.Controls.Add(_smtpPortNumeric, 1, 2)
            layout.Controls.Add(_emailFromLabel, 0, 3)
            layout.Controls.Add(_emailFromTextBox, 1, 3)
            layout.Controls.Add(_emailToLabel, 0, 4)
            layout.Controls.Add(_emailToTextBox, 1, 4)
            layout.Controls.Add(_useSSLCheckBox, 1, 5)
            layout.Controls.Add(_useCredentialsCheckBox, 0, 6)
            layout.Controls.Add(_credentialComboBox, 1, 6)
            layout.Controls.Add(_testEmailButton, 1, 7)

            emailTab.Controls.Add(layout)
            _tabControl.TabPages.Add(emailTab)
        End Sub

        Private Sub InitializeLoggingTab()
            Dim loggingTab As New TabPage("Logging")
            Dim layout As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .Padding = New Padding(10),
            .ColumnCount = 3,
            .RowCount = 5
        }

            ' Set column styles - adjusted for better proportions
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 30)) ' Labels
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 55)) ' Controls
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 15)) ' Browse button

            ' Initialize labels
            _logFolderLabel = New Label With {
                .Text = "Log Folder:",
                .Dock = DockStyle.Fill
                }

            _logLevelLabel = New Label With {
                .Text = "Log Level:",
                .Dock = DockStyle.Fill
                }

            _alertLevelLabel = New Label With {
                .Text = "Alert Level:",
                .Dock = DockStyle.Fill
                }

            _maxLogSizeLabel = New Label With {
                .Text = "Max Log Size (MB):",
                .Dock = DockStyle.Fill
                }

            _logRetentionLabel = New Label With {
                .Text = "Retention Days:",
                .Dock = DockStyle.Fill
                }

            ' Initialize controls
            _logFolderTextBox = New TextBox With {
            .Dock = DockStyle.Fill,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right
        }

            _browseLogFolderButton = New Button With {
            .Text = "Browse...",
            .Size = New Size(80, 23),
            .Anchor = AnchorStyles.Left
        }

            _logLevelComboBox = New ComboBox With {
            .Dock = DockStyle.Fill,
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
            _logLevelComboBox.Items.AddRange([Enum].GetNames(GetType(Logger.LogLevel)))

            _maxLogSizeNumeric = New NumericUpDown With {
            .Minimum = 1,
            .Maximum = 1000,
            .Value = 10,
            .Dock = DockStyle.Fill
        }

            _logRetentionNumeric = New NumericUpDown With {
            .Minimum = 1,
            .Maximum = 365,
            .Value = 30,
            .Dock = DockStyle.Fill
        }

            _alertLevelComboBox = New ComboBox With {
            .Dock = DockStyle.Fill,
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
            _alertLevelComboBox.Items.AddRange([Enum].GetNames(GetType(Logger.LogLevel)))

            ' Add controls to layout
            _logFolderLabel = New Label With {
            .Text = "Log Folder:",
            .Dock = DockStyle.Fill
            }

            layout.Controls.Add(_logFolderLabel, 0, 0)
            layout.Controls.Add(_logFolderTextBox, 1, 0)
            layout.Controls.Add(_browseLogFolderButton, 2, 0)
            layout.Controls.Add(_logLevelLabel, 0, 1)
            layout.Controls.Add(_logLevelComboBox, 1, 1)
            layout.Controls.Add(_alertLevelLabel, 0, 2)
            layout.Controls.Add(_alertLevelComboBox, 1, 2)
            layout.Controls.Add(_maxLogSizeLabel, 0, 3)
            layout.Controls.Add(_maxLogSizeNumeric, 1, 3)
            layout.Controls.Add(_logRetentionLabel, 0, 4)
            layout.Controls.Add(_logRetentionNumeric, 1, 4)

            ' Add browse folder handler
            AddHandler _browseLogFolderButton.Click, AddressOf BrowseLogFolder_Click

            loggingTab.Controls.Add(layout)
            _tabControl.TabPages.Add(loggingTab)
        End Sub

        Private Sub InitializeSqlTab()
            _sqlTab = New TabPage(TranslationManager.Instance.GetTranslation("SqlTab"))
            Dim layout As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .Padding = New Padding(10),
            .ColumnCount = 2,
            .RowCount = 4
        }

            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 30))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 70))

            _defaultServerLabel = New Label With {
            .Text = "Default Server:",
            .Dock = DockStyle.Fill
        }

            _defaultDatabaseLabel = New Label With {
            .Text = "Default Database:",
            .Dock = DockStyle.Fill
        }

            _commandTimeoutLabel = New Label With {
            .Text = "Command Timeout:",
            .Dock = DockStyle.Fill
        }

            _maxBatchSizeLabel = New Label With {
            .Text = "Max Batch Size:",
            .Dock = DockStyle.Fill
        }

            _defaultServerTextBox = New TextBox With {
            .Dock = DockStyle.Fill
        }

            _defaultDatabaseTextBox = New TextBox With {
            .Dock = DockStyle.Fill
        }

            _commandTimeoutNumeric = New NumericUpDown With {
            .Minimum = 30,
            .Maximum = 3600,
            .Value = 300,
            .Dock = DockStyle.Fill
        }

            _maxBatchSizeNumeric = New NumericUpDown With {
            .Minimum = 100,
            .Maximum = 10000,
            .Value = 1000,
            .Dock = DockStyle.Fill
        }

            ' Add tooltips
            Dim toolTip As New ToolTip()
            toolTip.SetToolTip(_defaultServerTextBox, TranslationManager.Instance.GetTranslation("SqlTab.Tooltip.Server"))
            toolTip.SetToolTip(_defaultDatabaseTextBox, TranslationManager.Instance.GetTranslation("SqlTab.Tooltip.Database"))
            toolTip.SetToolTip(_commandTimeoutNumeric, TranslationManager.Instance.GetTranslation("SqlTab.Tooltip.Timeout"))
            toolTip.SetToolTip(_maxBatchSizeNumeric, TranslationManager.Instance.GetTranslation("SqlTab.Tooltip.BatchSize"))

            ' Add controls to layout
            layout.Controls.Add(_defaultServerLabel, 0, 0)
            layout.Controls.Add(_defaultServerTextBox, 1, 0)
            layout.Controls.Add(_defaultDatabaseLabel, 0, 1)
            layout.Controls.Add(_defaultDatabaseTextBox, 1, 1)
            layout.Controls.Add(_commandTimeoutLabel, 0, 2)
            layout.Controls.Add(_commandTimeoutNumeric, 1, 2)
            layout.Controls.Add(_maxBatchSizeLabel, 0, 3)
            layout.Controls.Add(_maxBatchSizeNumeric, 1, 3)

            _sqlTab.Controls.Add(layout)
            _tabControl.TabPages.Add(_sqlTab)
        End Sub

        Private Function IsValidEmail(email As String) As Boolean
            Try
                Dim addr = New System.Net.Mail.MailAddress(email)
                Return addr.Address = email
            Catch
                Return False
            End Try
        End Function

        Private Sub LoadSettings()
            Try
                ' Load logging settings
                Dim logSettings = _xmlManager.GetLogSettings()
                _logFolderTextBox.Text = GetSettingValue(logSettings, "LogFolder", Path.Combine(Application.StartupPath, "LiteTaskData", "logs"))
                _logLevelComboBox.SelectedItem = GetSettingValue(logSettings, "LogLevel", "Info")
                _maxLogSizeNumeric.Value = Integer.Parse(GetSettingValue(logSettings, "MaxLogSize", "10"))
                _logRetentionNumeric.Value = Integer.Parse(GetSettingValue(logSettings, "LogRetentionDays", "30"))
                _alertLevelComboBox.SelectedItem = GetSettingValue(logSettings, "AlertLevel", "Error")

                ' Load email settings
                Dim emailSettings = _xmlManager.GetEmailSettings()
                _enableNotificationsCheckBox.Checked = Boolean.Parse(GetSettingValue(emailSettings, "NotificationsEnabled", "False"))
                _smtpServerTextBox.Text = GetSettingValue(emailSettings, "SmtpServer", "")
                _smtpPortNumeric.Value = Integer.Parse(GetSettingValue(emailSettings, "SmtpPort", "25"))
                _useSSLCheckBox.Checked = Boolean.Parse(GetSettingValue(emailSettings, "UseSSL", "True"))
                _emailFromTextBox.Text = GetSettingValue(emailSettings, "EmailFrom", "")
                _emailToTextBox.Text = GetSettingValue(emailSettings, "EmailTo", "")

                ' Load credential settings if they exist
                If emailSettings.ContainsKey("UseCredentials") Then
                    _useCredentialsCheckBox.Checked = Boolean.Parse(GetSettingValue(emailSettings, "UseCredentials", "False"))
                    If emailSettings.ContainsKey("CredentialTarget") Then
                        Dim credTarget = emailSettings("CredentialTarget")
                        Dim index = _credentialComboBox.Items.IndexOf(credTarget)
                        If index >= 0 Then
                            _credentialComboBox.SelectedIndex = index
                        End If
                    End If
                End If

                UpdateEmailControls(Nothing, EventArgs.Empty)

                ' Load SQL settings
                Dim sqlConfig = _xmlManager.GetSqlConfiguration()
                _defaultServerTextBox.Text = GetSettingValue(sqlConfig, "DefaultServer", "")
                _defaultDatabaseTextBox.Text = GetSettingValue(sqlConfig, "DefaultDatabase", "")
                _commandTimeoutNumeric.Value = Integer.Parse(GetSettingValue(sqlConfig, "CommandTimeout", "300"))
                _maxBatchSizeNumeric.Value = Integer.Parse(GetSettingValue(sqlConfig, "MaxBatchSize", "1000"))

            Catch ex As Exception
                _logger.LogError($"Error loading settings: {ex.Message}")
                MessageBox.Show($"Error loading settings: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

                ' Set default values if loading fails
                SetDefaultValues()
            End Try
        End Sub

        Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
            If DialogResult = DialogResult.OK Then
                e.Cancel = Not ValidateSettings()
            End If
            MyBase.OnFormClosing(e)
        End Sub

        Private Sub SaveSettings()
            Try
                If Not ValidateSettings() Then
                    DialogResult = DialogResult.None
                    Return
                End If

                ' Save logging settings
                _xmlManager.SaveLogSettings(
                    _logFolderTextBox.Text,
                    _logLevelComboBox.SelectedItem.ToString(),
                    CInt(_maxLogSizeNumeric.Value),
                    CInt(_logRetentionNumeric.Value))

                ' Save email settings
                _xmlManager.WriteValue("EmailSettings", "NotificationsEnabled", _enableNotificationsCheckBox.Checked.ToString())
                _xmlManager.WriteValue("EmailSettings", "SmtpServer", _smtpServerTextBox.Text)
                _xmlManager.WriteValue("EmailSettings", "SmtpPort", _smtpPortNumeric.Value.ToString())
                _xmlManager.WriteValue("EmailSettings", "UseSSL", _useSSLCheckBox.Checked.ToString())
                _xmlManager.WriteValue("EmailSettings", "EmailFrom", _emailFromTextBox.Text)
                _xmlManager.WriteValue("EmailSettings", "EmailTo", _emailToTextBox.Text)
                _xmlManager.WriteValue("EmailSettings", "AlertLevel", _alertLevelComboBox.SelectedItem.ToString())

                ' Save credential settings
                _xmlManager.WriteValue("EmailSettings", "UseCredentials", _useCredentialsCheckBox.Checked.ToString())
                _xmlManager.WriteValue("EmailSettings", "CredentialTarget",
                    If(_useCredentialsCheckBox.Checked AndAlso _credentialComboBox.SelectedIndex > 0,
                       _credentialComboBox.SelectedItem.ToString(),
                       ""))

                ' Save SQL settings
                _xmlManager.WriteValue("SqlConfiguration", "DefaultServer", _defaultServerTextBox.Text)
                _xmlManager.WriteValue("SqlConfiguration", "DefaultDatabase", _defaultDatabaseTextBox.Text)
                _xmlManager.WriteValue("SqlConfiguration", "CommandTimeout", _commandTimeoutNumeric.Value.ToString())
                _xmlManager.WriteValue("SqlConfiguration", "MaxBatchSize", _maxBatchSizeNumeric.Value.ToString())

                DialogResult = DialogResult.OK
                _logger.LogInfo("Settings saved successfully")
                MessageBox.Show("Settings saved successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Close()

            Catch ex As Exception
                _logger.LogError($"Error saving settings: {ex.Message}")
                MessageBox.Show($"Error saving settings: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                DialogResult = DialogResult.None
            End Try
        End Sub

        Private Sub SetDefaultValues()
            ' Set default logging values
            _logFolderTextBox.Text = Path.Combine(Application.StartupPath, "LiteTaskData", "logs")
            _logLevelComboBox.SelectedItem = "Info"
            _maxLogSizeNumeric.Value = 10
            _logRetentionNumeric.Value = 30
            _alertLevelComboBox.SelectedItem = "Error"

            ' Set default email values
            _enableNotificationsCheckBox.Checked = False
            _smtpServerTextBox.Text = ""
            _smtpPortNumeric.Value = 25
            _useSSLCheckBox.Checked = True
            _emailFromTextBox.Text = ""
            _emailToTextBox.Text = ""
            _useCredentialsCheckBox.Checked = False
            _credentialComboBox.SelectedIndex = 0
        End Sub

        Private Sub TestEmailButton_Click(sender As Object, e As EventArgs)
            Try
                If Not ValidateSettings() Then
                    Return
                End If

                Dim emailUtils = ApplicationContainer.GetService(Of EmailUtils)()
                Dim testMessage = "This is a test email from LiteTask."
                Dim subject = "LiteTask Test Email"

                Dim credential As CredentialInfo = Nothing
                If _useCredentialsCheckBox.Checked AndAlso _credentialComboBox.SelectedIndex > 0 Then
                    credential = _credentialManager.GetCredential(_credentialComboBox.SelectedItem.ToString(), "Windows Vault")
                End If

                ' Create test notification
                emailUtils.SendEmailReport(subject, testMessage)
                MessageBox.Show("Test email sent successfully.", "Test Email", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                _logger.LogError($"Error sending test email: {ex.Message}")
                MessageBox.Show($"Error sending test email: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub


        Private Sub UpdateEmailControls(sender As Object, e As EventArgs)
            Dim enabled = _enableNotificationsCheckBox.Checked
            _smtpServerTextBox.Enabled = enabled
            _smtpPortNumeric.Enabled = enabled
            _emailFromTextBox.Enabled = enabled
            _emailToTextBox.Enabled = enabled
            _useSSLCheckBox.Enabled = enabled
            _useCredentialsCheckBox.Enabled = enabled
            _credentialComboBox.Enabled = enabled AndAlso _useCredentialsCheckBox.Checked
            _testEmailButton.Enabled = enabled
        End Sub

        Private Sub UpdateCredentialControls(sender As Object, e As EventArgs)
            _credentialComboBox.Enabled = _useCredentialsCheckBox.Checked AndAlso _enableNotificationsCheckBox.Checked
        End Sub

        Private Function ValidateSettings() As Boolean
            ' Create validation context
            Dim emailSettings As New Dictionary(Of String, String) From {
        {"SmtpServer", _smtpServerTextBox.Text},
        {"SmtpPort", _smtpPortNumeric.Value.ToString()},
        {"EmailFrom", _emailFromTextBox.Text},
        {"EmailTo", _emailToTextBox.Text},
        {"NotificationsEnabled", _enableNotificationsCheckBox.Checked.ToString()}
    }

            Dim errors = _settingValidator.ValidateEmailSettings(emailSettings)
            errors.AddRange(_settingValidator.ValidateLogSettings(GetCurrentLogSettings()))

            If errors.Any() Then
                MessageBox.Show(String.Join(Environment.NewLine, errors),
            "Validation Errors", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If
            Return True
        End Function

    End Class
End Namespace