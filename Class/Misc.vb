Imports System.Net.Mail

Namespace LiteTask

    Public Class CredentialInfo
        Public Property Target As String
        Public Property Username As String
        Public Property Password As String
        Public Property AccountType As String
        Public Property SecurePassword As SecureString
        Public Property RequiresElevation As Boolean

        Public Sub New()
            RequiresElevation = False
        End Sub


    End Class
    '-------------------------------------------------------------------------------
    Public Class EmailUtils
        Private ReadOnly _logger As Logger
        Private ReadOnly _xmlManager As XMLManager

        Public Sub New(logger As Logger, xmlManager As XMLManager)
            _logger = logger
            _xmlManager = xmlManager
        End Sub

        Public Sub SendEmailReport(subject As String, body As String)
            Try
                Dim emailSettings = _xmlManager.GetEmailSettings()
                Dim smtpServer = emailSettings("SmtpServer")
                Dim smtpPort = Integer.Parse(emailSettings("SmtpPort"))
                Dim emailFrom = emailSettings("EmailFrom")
                Dim emailTo = emailSettings("EmailTo")

                Using mail As New MailMessage()
                    mail.From = New MailAddress(emailFrom)
                    For Each recipient In emailTo.Split(";"c)
                        mail.To.Add(recipient.Trim())
                    Next
                    mail.Subject = subject
                    mail.Body = body

                    Using smtp As New SmtpClient(smtpServer, smtpPort)
                        smtp.Send(mail)
                    End Using
                End Using

                _logger.LogInfo($"Email report sent: {subject}")
            Catch ex As Exception
                _logger.LogError($"Failed to send email report: {ex.Message}")
                _logger.LogError($"StackTrace: {ex.StackTrace}")
            End Try
        End Sub
    End Class

    '-------------------------------------------------------------------------------
    Public Class ErrorNotifier
        Private ReadOnly _xmlManager As XMLManager
        Private ReadOnly _logger As Logger

        Public Sub New(xmlManager As XMLManager, logger As Logger)
            _xmlManager = xmlManager
            _logger = logger
        End Sub

        Public Sub NotifyError(errorMessage As String, errorLevel As String)
            Dim emailSettings = _xmlManager.GetEmailSettings()

            If Boolean.Parse(emailSettings("NotificationsEnabled")) AndAlso ShouldNotify(errorLevel, emailSettings("ErrorLevelThreshold")) Then
                SendErrorEmail(errorMessage, errorLevel, emailSettings)
            End If

            ' Always log the error
            _logger.LogError(errorMessage)
        End Sub

        Private Function ShouldNotify(errorLevel As String, threshold As String) As Boolean
            Dim levels = New List(Of String) From {"Info", "Warning", "Error"}
            Return levels.IndexOf(errorLevel) >= levels.IndexOf(threshold)
        End Function

        Private Sub SendErrorEmail(errorMessage As String, errorLevel As String, emailSettings As Dictionary(Of String, String))
            Try
                Using mail As New MailMessage()
                    mail.From = New MailAddress(emailSettings("EmailFrom"))
                    For Each recipient In emailSettings("EmailTo").Split(";"c)
                        mail.To.Add(recipient.Trim())
                    Next
                    mail.Subject = $"LiteTask Error Notification: {errorLevel}"
                    mail.Body = $"An error occurred in LiteTask:{Environment.NewLine}{Environment.NewLine}{errorMessage}"

                    Using smtp As New SmtpClient(emailSettings("SmtpServer"), Integer.Parse(emailSettings("SmtpPort")))
                        smtp.Send(mail)
                    End Using
                End Using
                _logger.LogInfo($"Error notification email sent successfully for {errorLevel} error")
            Catch ex As Exception
                _logger.LogError($"Failed to send error notification email: {ex.Message}")
            End Try
        End Sub
    End Class

    '-------------------------------------------------------------------------------
    Public Class FileTypeDetector
        Public Shared Function DetectTaskType(filePath As String) As ScheduledTask.TaskType
            Dim extension As String = Path.GetExtension(filePath).ToLower()
            Select Case extension
                Case ".ps1"
                    Return ScheduledTask.TaskType.PowerShell
                Case ".bat", ".cmd"
                    Return ScheduledTask.TaskType.Batch
                Case ".sql"
                    Return ScheduledTask.TaskType.SQL
                Case ".exe"
                    Return ScheduledTask.TaskType.Executable
                Case Else
                    ' Default to PowerShell for unknown extensions
                    Return ScheduledTask.TaskType.PowerShell
            End Select
        End Function
    End Class

    '-------------------------------------------------------------------------------
    Public Class LiteRunConfig
        Public Property Timeout As Integer
        Public Property LogOutputPath As String

        Public Sub New(defaults As Dictionary(Of String, String))
            Timeout = Integer.Parse(defaults("Timeout"))
            LogOutputPath = Path.Combine(Application.StartupPath, "LiteTaskData", "logs")
        End Sub

        Public Function GetCommandLineArguments() As String
            Dim args As New List(Of String)
            args.Add($"-to {Timeout}")
            args.Add($"-lo ""{Path.Combine(LogOutputPath, $"Run_{DateTime.Now:yyyyMMddHHmmss}.log")}""")
            Return String.Join(" ", args)
        End Function
    End Class

    '-------------------------------------------------------------------------------
    Public Enum LiteRunExitCode
        Success = 0
        InternalError = -1
        CommandLineError = -2
        FailedToLaunchApp = -3
        FailedToCopyLiteRun = -4
        ConnectionTimeout = -5
        ServiceInstallationFailed = -6
        ServiceCommunicationFailed = -7
        FailedToCopyApp = -8
        FailedToLaunchRemoteApp = -9
        AppTerminatedAfterTimeout = -10
        ForciblyStoppedByUser = -11
    End Enum

    Public Class PathConfiguration
        Public Property AppDataPath As String
        Public Property ToolsPath As String
        Public Property TempPath As String
        Public Property LogsPath As String
        Public Property SettingsPath As String

        Public Sub New()
            ' Default constructor required for DI
        End Sub
    End Class
    '-------------------------------------------------------------------------------
    Public Class PowerShellPathManager
        Private ReadOnly _logger As Logger
        Private ReadOnly _modulesPath As String

        Public Sub New(logger As Logger)
            _logger = logger
            _modulesPath = Path.Combine(Application.StartupPath, "LiteTaskData", "modules")
        End Sub

        Public Function GetModulePath() As String
            Return _modulesPath
        End Function

        Public Function CreateInitializationScript() As String
            ' This script will be prepended to all PowerShell executions
            Return $"
                $env:PSModulePath = '{_modulesPath}' + [System.IO.Path]::PathSeparator + $env:PSModulePath
                $ErrorActionPreference = 'Stop'
                Import-Module PowerShellGet -Force -ErrorAction SilentlyContinue
            "
        End Function

        Public Sub EnsureModulePathExists()
            If Not Directory.Exists(_modulesPath) Then
                Try
                    Directory.CreateDirectory(_modulesPath)
                    _logger.LogInfo($"Created PowerShell modules directory: {_modulesPath}")
                Catch ex As Exception
                    _logger.LogError($"Error creating modules directory: {ex.Message}")
                    Throw
                End Try
            End If
        End Sub

        Public Function CreatePowerShellInstance() As PowerShell
            Try
                ' Explicitly declare the type
                Dim initialSessionState As InitialSessionState = InitialSessionState.CreateDefault2()
                initialSessionState.ExecutionPolicy = ExecutionPolicy.Bypass

                ' Add our modules path to the session
                initialSessionState.ImportPSModule(Directory.GetDirectories(_modulesPath))

                Dim ps = PowerShell.Create(initialSessionState)
                ps.AddScript(CreateInitializationScript())
                Return ps
            Catch ex As Exception
                _logger.LogError($"Error creating PowerShell instance: {ex.Message}")
                Throw
            End Try
        End Function
    End Class

    Public Class TaskState
        Public Property IsRunning As Boolean
        Public Property LastStartTime As DateTime
        Public Property LastEndTime As DateTime?
        Public Property CurrentProcess As Integer?
        Public Property StatusMessage As String
        Public Property LastError As String
    End Class

    Public Class TaskCompletedEventArgs
        Inherits EventArgs

        Public Property Task As ScheduledTask

        Public Sub New(task As ScheduledTask)
            Me.Task = task
        End Sub
    End Class

    Public Class ProcessResult
        Public Property ExitCode As Integer
        Public Property Output As String
        Public Property GotError As String
    End Class

End Namespace