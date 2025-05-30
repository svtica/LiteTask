Imports System.Drawing

Namespace LiteTask
    Public Class ApplicationContext
        Inherits System.Windows.Forms.ApplicationContext

        Private ReadOnly _notifyIcon As NotifyIcon
        Private ReadOnly _mainForm As MainForm
        Private ReadOnly _logger As Logger
        Private ReadOnly _customScheduler As CustomScheduler
        Private ReadOnly _xmlManager As XMLManager
        Private _isQuitting As Boolean = False

        ' Define paths  
        Dim appDataPath = Path.Combine(Application.StartupPath, "LiteTaskData")
        Dim toolsPath = Path.Combine(appDataPath, "tools")
        Dim tempPath = Path.Combine(appDataPath, "temp")
        Dim logsPath = Path.Combine(appDataPath, "logs")
        Dim settingsPath = Path.Combine(appDataPath, "settings.xml")
        Dim defaultLogFilePath = Path.Combine(logsPath, "app_log.txt")

        Public Sub New()
            Try
                ' Ensure directories exist  
                Directory.CreateDirectory(appDataPath)
                Directory.CreateDirectory(toolsPath)
                Directory.CreateDirectory(tempPath)
                Directory.CreateDirectory(logsPath)

                ' Initialize services  
                _logger = ApplicationContainer.GetService(Of Logger)()
                _customScheduler = ApplicationContainer.GetService(Of CustomScheduler)()
                _xmlManager = ApplicationContainer.GetService(Of XMLManager)()

                ' Create the notify icon  
                _notifyIcon = New NotifyIcon() With {
            .Icon = New Icon(Path.Combine(Application.StartupPath, "res", "ico", "logo.ico")),
            .Visible = True,
            .Text = TranslationManager.Instance.GetTranslation("NotifyIcon.Text")
        }

                ' Create context menu  
                Dim contextMenu As New ContextMenuStrip()

                ' Add menu items with translations
                Dim showItem As New ToolStripMenuItem(
            TranslationManager.Instance.GetTranslation("Show"),
            Nothing,
            AddressOf ShowMainForm)

                Dim optionsItem As New ToolStripMenuItem(
            TranslationManager.Instance.GetTranslation("_optionsMenuItem.Text"),
            Nothing,
            AddressOf ShowSettings)

                Dim exitItem As New ToolStripMenuItem(
            TranslationManager.Instance.GetTranslation("Exit"),
            Nothing,
            AddressOf ExitApplication)

                contextMenu.Items.AddRange({showItem, optionsItem, New ToolStripSeparator(), exitItem})
                _notifyIcon.ContextMenuStrip = contextMenu

                ' Create main form but don't show it yet  
                _mainForm = New MainForm()
                AddHandler _mainForm.FormClosing, AddressOf MainForm_FormClosing

                ' Double click on tray icon shows main form  
                AddHandler _notifyIcon.MouseDoubleClick, AddressOf NotifyIcon_MouseDoubleClick

                ' Show loading form  
                Using loadingForm As New Loading()
                    loadingForm.ShowDialog()
                End Using

                ' Load initial settings  
                LoadSettings()

                'Show main form
                _mainForm.Show()

                _logger.LogInfo("Application context initialized successfully")
            Catch ex As Exception
                MessageBox.Show($"Error initializing application: {ex.Message}", "Initialization Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                _logger?.LogError($"Error in LiteTaskApplicationContext constructor: {ex.Message}")
                _logger?.LogError($"StackTrace: {ex.StackTrace}")
                Application.Exit()
            End Try
        End Sub

        Private Sub LoadSettings()
            Try
                ' Load log settings first
                Dim logSettings = _xmlManager.GetLogSettings()
                _logger.SetLogLevel(DirectCast([Enum].Parse(GetType(Logger.LogLevel), logSettings("LogLevel")), Logger.LogLevel))
                _logger.SetLogFolder(logSettings("LogFolder"))

                ' Then load email settings
                Dim emailSettings = _xmlManager.GetEmailSettings()
                If Boolean.Parse(emailSettings("NotificationsEnabled")) Then
                    ' Initialize email notification system
                End If

                _logger.LogInfo("Settings loaded successfully")
            Catch ex As Exception
                _logger.LogError($"Error loading settings: {ex.Message}")
                MessageBox.Show($"Error loading settings: {ex.Message}", "Settings Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Sub

        Private Sub ShowMainForm(sender As Object, e As EventArgs)
            If _mainForm.IsDisposed Then
                Dim mainForm As New MainForm()
                Application.Run(mainForm)
                AddHandler _mainForm.FormClosing, AddressOf MainForm_FormClosing
            End If

            If Not _mainForm.Visible Then
                _mainForm.Show()
                _mainForm.WindowState = FormWindowState.Normal
                _mainForm.BringToFront()
            End If
        End Sub

        Private Sub ShowSettings(sender As Object, e As EventArgs)
            Using optionsForm As New OptionsForm(_xmlManager)
                If optionsForm.ShowDialog() = DialogResult.OK Then
                    LoadSettings()
                End If
            End Using
        End Sub

        Private Sub ExitApplication(sender As Object, e As EventArgs)
            Try
                ' Ask for confirmation before exiting with translated message 
                If MessageBox.Show(
            TranslationManager.Instance.GetTranslation("ConfirmExit"),
            TranslationManager.Instance.GetTranslation("ConfirmExitTitle"),
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question) = DialogResult.Yes Then

                    _isQuitting = True ' Set flag to indicate we're actually quitting

                    ' Save any pending changes  
                    _customScheduler.SaveTasks()

                    ' Clean up resources  
                    If _mainForm IsNot Nothing AndAlso Not _mainForm.IsDisposed Then
                        _mainForm.Close()
                        _mainForm.Dispose()
                    End If

                    If _notifyIcon IsNot Nothing Then
                        _notifyIcon.Visible = False
                        _notifyIcon.Dispose()
                    End If

                    ' Exit the application  
                    Application.Exit()
                End If
            Catch ex As Exception
                _logger.LogError($"Error during application exit: {ex.Message}")
                _logger.LogError($"StackTrace: {ex.StackTrace}")
                Application.Exit()
            End Try
        End Sub

        Private Sub NotifyIcon_MouseDoubleClick(sender As Object, e As MouseEventArgs)
            If e.Button = MouseButtons.Left Then
                ShowMainForm(sender, e)
            End If
        End Sub

        Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs)
            If e.CloseReason = CloseReason.UserClosing AndAlso Not _isQuitting Then
                e.Cancel = True
                _mainForm.Hide()
                _notifyIcon.ShowBalloonTip(
            3000,
            TranslationManager.Instance.GetTranslation("NotifyIcon.BalloonTipTitle"),
            TranslationManager.Instance.GetTranslation("NotifyIcon.BalloonTipText"),
            ToolTipIcon.Info)
            End If
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            Try
                If disposing Then
                    If _mainForm IsNot Nothing AndAlso Not _mainForm.IsDisposed Then
                        _mainForm.Dispose()
                    End If
                    If _notifyIcon IsNot Nothing Then
                        _notifyIcon.Dispose()
                    End If
                End If
            Finally
                MyBase.Dispose(disposing)
            End Try
        End Sub

    End Class
End Namespace