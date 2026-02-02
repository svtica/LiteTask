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

        ''' <summary>
        ''' Gets all standard PowerShell module paths from the system.
        ''' </summary>
        Public Function GetSystemModulePaths() As List(Of String)
            Dim paths As New List(Of String)

            Try
                ' Add the LiteTask local modules path first
                paths.Add(_modulesPath)

                ' Standard PowerShell module locations
                Dim programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
                Dim userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
                Dim systemRoot = Environment.GetEnvironmentVariable("SystemRoot")

                ' PowerShell 7+ module paths
                Dim ps7ModulePath = Path.Combine(programFiles, "PowerShell", "Modules")
                If Directory.Exists(ps7ModulePath) Then paths.Add(ps7ModulePath)

                Dim ps7PreviewModulePath = Path.Combine(programFiles, "PowerShell", "7", "Modules")
                If Directory.Exists(ps7PreviewModulePath) Then paths.Add(ps7PreviewModulePath)

                ' Windows PowerShell 5.1 module paths
                If systemRoot IsNot Nothing Then
                    Dim winPsModulePath = Path.Combine(systemRoot, "System32", "WindowsPowerShell", "v1.0", "Modules")
                    If Directory.Exists(winPsModulePath) Then paths.Add(winPsModulePath)
                End If

                ' Program Files module path (shared between PS editions)
                Dim pfModulePath = Path.Combine(programFiles, "WindowsPowerShell", "Modules")
                If Directory.Exists(pfModulePath) Then paths.Add(pfModulePath)

                ' User-scoped module paths
                Dim userPsModulePath = Path.Combine(userProfile, "Documents", "PowerShell", "Modules")
                If Directory.Exists(userPsModulePath) Then paths.Add(userPsModulePath)

                Dim userWinPsModulePath = Path.Combine(userProfile, "Documents", "WindowsPowerShell", "Modules")
                If Directory.Exists(userWinPsModulePath) Then paths.Add(userWinPsModulePath)

                ' Also check the current PSModulePath environment variable for any custom paths
                Dim envModulePath = Environment.GetEnvironmentVariable("PSModulePath")
                If Not String.IsNullOrEmpty(envModulePath) Then
                    For Each envPath In envModulePath.Split(Path.PathSeparator)
                        Dim trimmedPath = envPath.Trim()
                        If Not String.IsNullOrEmpty(trimmedPath) AndAlso Directory.Exists(trimmedPath) AndAlso Not paths.Contains(trimmedPath) Then
                            paths.Add(trimmedPath)
                        End If
                    Next
                End If

            Catch ex As Exception
                _logger.LogError($"Error discovering system module paths: {ex.Message}")
            End Try

            Return paths
        End Function

        Public Function CreateInitializationScript() As String
            ' Build a comprehensive PSModulePath that includes all system paths
            Dim allPaths = GetSystemModulePaths()
            Dim pathString = String.Join([Char].ToString(Path.PathSeparator), allPaths)

            ' This script will be prepended to all PowerShell executions
            Return $"
                $env:PSModulePath = '{pathString}' + [System.IO.Path]::PathSeparator + $env:PSModulePath
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

        ''' <summary>
        ''' Checks whether a PowerShell module is available in any known module path.
        ''' Returns the path where the module was found, or Nothing if not found.
        ''' </summary>
        Public Function FindModule(moduleName As String) As String
            Try
                For Each basePath In GetSystemModulePaths()
                    If Not Directory.Exists(basePath) Then Continue For

                    ' Check for module directory (most common layout)
                    Dim moduleDirPath = Path.Combine(basePath, moduleName)
                    If Directory.Exists(moduleDirPath) Then
                        ' Verify it contains a module manifest or script module
                        If File.Exists(Path.Combine(moduleDirPath, $"{moduleName}.psd1")) OrElse
                           File.Exists(Path.Combine(moduleDirPath, $"{moduleName}.psm1")) Then
                            Return moduleDirPath
                        End If

                        ' Check versioned subdirectories (e.g., ImportExcel/7.8.6/ImportExcel.psd1)
                        For Each versionDir In Directory.GetDirectories(moduleDirPath)
                            If File.Exists(Path.Combine(versionDir, $"{moduleName}.psd1")) OrElse
                               File.Exists(Path.Combine(versionDir, $"{moduleName}.psm1")) Then
                                Return versionDir
                            End If
                        Next
                    End If
                Next
            Catch ex As Exception
                _logger.LogError($"Error searching for module '{moduleName}': {ex.Message}")
            End Try

            Return Nothing
        End Function

        ''' <summary>
        ''' Verifies a list of module names and logs their availability.
        ''' Returns a list of modules that could not be found.
        ''' </summary>
        Public Function VerifyModules(moduleNames As IEnumerable(Of String)) As List(Of String)
            Dim missingModules As New List(Of String)

            For Each moduleName In moduleNames
                Dim foundPath = FindModule(moduleName)
                If foundPath IsNot Nothing Then
                    _logger.LogInfo($"PowerShell module '{moduleName}' found at: {foundPath}")
                Else
                    _logger.LogWarning($"PowerShell module '{moduleName}' not found in any module path")
                    missingModules.Add(moduleName)
                End If
            Next

            Return missingModules
        End Function

        ''' <summary>
        ''' Scans a PowerShell script file for required modules by looking for
        ''' #Requires -Module and Import-Module statements.
        ''' </summary>
        Public Shared Function DetectRequiredModules(scriptContent As String) As List(Of String)
            Dim modules As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            Try
                ' Match #Requires -Module ModuleName or #Requires -Modules ModuleName, Module2
                Dim requiresPattern As New Text.RegularExpressions.Regex(
                    "#Requires\s+-Modules?\s+(.+)",
                    Text.RegularExpressions.RegexOptions.IgnoreCase Or Text.RegularExpressions.RegexOptions.Multiline)

                For Each match As Text.RegularExpressions.Match In requiresPattern.Matches(scriptContent)
                    Dim moduleList = match.Groups(1).Value
                    For Each modName In moduleList.Split(","c)
                        Dim cleaned = modName.Trim().Trim("'"c, """"c)
                        ' Handle @{ModuleName = 'Name'; ...} hashtable syntax
                        If cleaned.StartsWith("@{") Then
                            Dim nameMatch = Text.RegularExpressions.Regex.Match(cleaned, "ModuleName\s*=\s*['""]?(\w+)")
                            If nameMatch.Success Then modules.Add(nameMatch.Groups(1).Value)
                        ElseIf Not String.IsNullOrWhiteSpace(cleaned) Then
                            modules.Add(cleaned)
                        End If
                    Next
                Next

                ' Match Import-Module ModuleName (not variables like $module)
                Dim importPattern As New Text.RegularExpressions.Regex(
                    "Import-Module\s+(?:-Name\s+)?['""]?([A-Za-z][\w.]+)['""]?",
                    Text.RegularExpressions.RegexOptions.IgnoreCase Or Text.RegularExpressions.RegexOptions.Multiline)

                For Each match As Text.RegularExpressions.Match In importPattern.Matches(scriptContent)
                    Dim modName = match.Groups(1).Value.Trim()
                    ' Skip common built-in modules that don't need verification
                    If Not String.IsNullOrWhiteSpace(modName) AndAlso
                       Not modName.Equals("Microsoft.PowerShell.Management", StringComparison.OrdinalIgnoreCase) AndAlso
                       Not modName.Equals("Microsoft.PowerShell.Utility", StringComparison.OrdinalIgnoreCase) Then
                        modules.Add(modName)
                    End If
                Next

            Catch ex As Exception
                ' Don't fail script execution if detection fails
            End Try

            Return modules.ToList()
        End Function

        Public Function CreatePowerShellInstance() As PowerShell
            Try
                ' Explicitly declare the type
                Dim initialSessionState As InitialSessionState = InitialSessionState.CreateDefault2()
                initialSessionState.ExecutionPolicy = ExecutionPolicy.Bypass

                ' Only import modules from the local LiteTask modules directory (safe, controlled)
                ' System module paths are added to $env:PSModulePath via the initialization script
                ' so PowerShell's auto-loading can find them on demand without forcing eager loading
                Dim localModuleDirs = If(Directory.Exists(_modulesPath),
                                         Directory.GetDirectories(_modulesPath),
                                         Array.Empty(Of String)())
                If localModuleDirs.Length > 0 Then initialSessionState.ImportPSModule(localModuleDirs)

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

    '-------------------------------------------------------------------------------

    ''' Safe wrapper for mutex operations
    Public Class MutexHandle
        Implements IDisposable

        Private ReadOnly _mutexName As String
        Private ReadOnly _taskName As String
        Private ReadOnly _logger As Logger
        Private _mutex As Mutex
        Private _hasLock As Boolean
        Private _disposed As Boolean

        Public Sub New(mutexName As String, taskName As String, logger As Logger)
            _mutexName = mutexName
            _taskName = taskName
            _logger = logger
        End Sub

        Public Function TryAcquire(timeoutSeconds As Integer) As Boolean
            Try
                _mutex = New Mutex(False, _mutexName)
                _hasLock = _mutex.WaitOne(TimeSpan.FromSeconds(timeoutSeconds))
                Return _hasLock
                
            Catch ex As AbandonedMutexException
                _logger.LogWarning($"Recovered from abandoned mutex for task: {_taskName}")
                _hasLock = True
                Return True
            Catch ex As Exception
                _logger.LogError($"Error acquiring mutex for {_taskName}: {ex.Message}")
                Return False
            End Try
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            If Not _disposed Then
                Try
                    If _hasLock AndAlso _mutex IsNot Nothing Then
                        _mutex.ReleaseMutex()
                        _logger.LogInfo($"Mutex released for task: {_taskName}")
                    End If
                Catch ex As Exception
                    _logger.LogError($"Error releasing mutex for {_taskName}: {ex.Message}")
                Finally
                    _mutex?.Dispose()
                    _disposed = True
                End Try
            End If
        End Sub
    End Class

    '-------------------------------------------------------------------------------

    ''' Information about an active mutex
    Public Class MutexInfo
        Public Property TaskName As String
        Public Property AcquiredAt As DateTime
        Public Property MutexHandle As MutexHandle
    End Class

    '-------------------------------------------------------------------------------

    ''' Diagnostic information for mutex monitoring
    Public Class MutexDiagnosticInfo
        Public Property TaskName As String
        Public Property AcquiredAt As DateTime
        Public Property DurationMinutes As Double
        Public Property IsStale As Boolean
        
        Public Overrides Function ToString() As String
            Return $"Task: {TaskName}, Duration: {DurationMinutes:F1}min, Stale: {IsStale}"
        End Function
    End Class

End Namespace