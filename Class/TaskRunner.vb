Imports System.Collections.ObjectModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel

Namespace LiteTask
    Public Class TaskRunner

        ' TODO: Phase 2 Security Improvements
        ' - Implement audit logging for all database operations

        Private ReadOnly _logger As Logger
        Private ReadOnly _credentialManager As CredentialManager
        Private ReadOnly _runspace As Runspace
        Private ReadOnly _liteRunConfig As LiteRunConfig
        Private ReadOnly _toolManager As ToolManager
        Private ReadOnly _xmlManager As XMLManager
        Private ReadOnly _logPath As String
        Private ReadOnly _powerShellPathManager As PowerShellPathManager
        Private ReadOnly _sqlTab As SqlTab
        Private ReadOnly _sanitizer As New SqlCommandSanitizer()
        Private ReadOnly _embeddedPsExec As Byte()
        Private _connectionStringBase As String
        Private _sqlConfig As Dictionary(Of String, String)
        Public Event OutputReceived(sender As Object, data As String)
        Public Event ErrorReceived(sender As Object, data As String)

        Public Sub New(logger As Logger, credentialManager As CredentialManager, toolManager As ToolManager, logPath As String, XMLManager As XMLManager)
            If logger Is Nothing Then
                Throw New ArgumentNullException(NameOf(logger))
            End If
            If credentialManager Is Nothing Then
                Throw New ArgumentNullException(NameOf(credentialManager))
            End If
            If toolManager Is Nothing Then
                Throw New ArgumentNullException(NameOf(toolManager))
            End If
            If String.IsNullOrEmpty(logPath) Then
                Throw New ArgumentNullException(NameOf(logPath))
            End If
            If XMLManager Is Nothing Then
                Throw New ArgumentNullException(NameOf(XMLManager))
            End If

            _logger = logger
            _credentialManager = credentialManager
            _toolManager = toolManager
            _logPath = logPath
            _xmlManager = XMLManager
            _powerShellPathManager = New PowerShellPathManager(logger)
            '_xmlManager = ApplicationContainer.GetService(Of XMLManager)()
            _sqlConfig = _xmlManager.GetSqlConfiguration()

            ' Ensure required directories exist
            Directory.CreateDirectory(_logPath)
            Directory.CreateDirectory(_toolManager._toolsPath)

            ' Check if embedded resources are available
            'Dim assembly = GetType(TaskRunner).Assembly
            'For Each name In assembly.GetManifestResourceNames()
            '    _logger.LogInfo($"Found resource: {name}")
            'Next

            ' Verify required tools
            VerifyRequiredTools()

            ' Initialize PowerShell environment
            _powerShellPathManager.EnsureModulePathExists()
            Dim initialSessionState As InitialSessionState = InitialSessionState.CreateDefault2()
            initialSessionState.ExecutionPolicy = ExecutionPolicy.Bypass

            ' Only import modules from the local LiteTask modules directory (safe, controlled)
            ' System module paths are added to $env:PSModulePath via the initialization script
            ' so PowerShell's auto-loading can find them on demand without forcing eager loading
            Dim localModulePath = _powerShellPathManager.GetModulePath()
            If Directory.Exists(localModulePath) Then
                Dim localModuleDirs = Directory.GetDirectories(localModulePath)
                If localModuleDirs.Length > 0 Then
                    initialSessionState.ImportPSModule(localModuleDirs)
                End If
            End If

            _runspace = RunspaceFactory.CreateRunspace(initialSessionState)
            _runspace.Open()

            ' Finally load embedded PsExec
            GetEmbeddedPsExec()

            ' Clean up any old temporary files from previous runs
            CleanupOldTempFiles()

        End Sub

        'Private Function DataTableToString(dt As DataTable) As String
        '    ' TODO: Phase 2 Security Improvements
        '    ' - Add data masking for sensitive columns
        '    ' - Implement output size limits
        '    ' - Add data type validation
        '    ' - Remove potentially dangerous content from output

        '    Try
        '        Dim result As New StringBuilder()

        '        ' Add headers with basic sanitization
        '        result.AppendLine(String.Join(vbTab,
        '        dt.Columns.Cast(Of DataColumn)().
        '        Select(Function(c) _sanitizer.SanitizeColumnName(c.ColumnName))))

        '        ' Add rows with basic sanitization
        '        For Each row As DataRow In dt.Rows
        '            result.AppendLine(String.Join(vbTab,
        '            row.ItemArray.Select(Function(item) _sanitizer.SanitizeOutput(
        '                        If(item Is Nothing OrElse item Is DBNull.Value, "[NULL]", item.ToString())))))
        '        Next

        '        Return result.ToString()
        '    Catch ex As Exception
        '        _logger.LogError($"Error converting DataTable to string: {ex.Message}")
        '        Return String.Empty
        '    End Try
        'End Function

        Public Sub Dispose()
            If _runspace IsNot Nothing Then
                _runspace.Dispose()
            End If
        End Sub

        Private Function EscapeArgument(argument As String) As String
            If String.IsNullOrEmpty(argument) Then
                Return ""
            End If

            ' Check if argument needs quotes
            If Not argument.Contains(" ") AndAlso Not argument.Contains("\t") Then
                Return argument
            End If

            ' Escape by quotes and handle embedded quotes
            Return $"""{argument.Replace("""", "\""")}"""
        End Function

        Public Async Function ExecuteBatchTask(taskAction As TaskAction, credential As CredentialInfo) As Task(Of Boolean)
            Try
                _logger.LogInfo($"Executing Batch task: {taskAction.Name}")

                If credential IsNot Nothing Then
                    ' Parse UNC paths from batch content
                    Dim batchContent = File.ReadAllText(taskAction.Target)
                    Dim uncPaths = ParseUNCPaths(batchContent)
                    Dim envPassVar = $"BATCH_PASS_{Guid.NewGuid().ToString("N")}"

                    Try
                        ' Set up credentials
                        Using securePass As New SecureString()
                            For Each c In New NetworkCredential("", credential.SecurePassword).Password
                                securePass.AppendChar(c)
                            Next
                            securePass.MakeReadOnly()
                            Environment.SetEnvironmentVariable(envPassVar,
                        New NetworkCredential("", securePass).Password,
                        EnvironmentVariableTarget.Process)
                        End Using

                        ' Handle domain credentials
                        Dim baseArgs = "-accepteula -nobanner -i"
                        If taskAction.RequiresElevation Then baseArgs &= " -h"

                        If credential.Username.Contains("\") Then
                            Dim parts = credential.Username.Split("\"c)
                            baseArgs &= $" -u {parts(0)}\{parts(1)}"
                        Else
                            baseArgs &= $" -u {credential.Username}"
                        End If

                        ' Build command with network path authentication
                        Dim modifiedContent = New StringBuilder("@echo off" & Environment.NewLine)
                        For Each uncPath In uncPaths.Distinct()
                            modifiedContent.AppendLine($"net use {uncPath} /user:{credential.Username} %{envPassVar}%")
                        Next
                        modifiedContent.AppendLine(batchContent)
                        For Each uncPath In uncPaths.Distinct()
                            modifiedContent.AppendLine($"net use {uncPath} /delete")
                        Next

                        ' Execute using PsExec
                        Using ms As New MemoryStream()
                            Using writer As New StreamWriter(ms)
                                Await writer.WriteAsync(modifiedContent.ToString())
                                Await writer.FlushAsync()
                                ms.Position = 0

                                Dim psExecPath = Path.Combine(_toolManager._toolsPath, "PsExec64.exe")
                                Dim fullCommand = $"{baseArgs} -p ""%{envPassVar}%"" cmd.exe /c -"

                                Using process As New Process With {
                            .StartInfo = New ProcessStartInfo With {
                                .FileName = psExecPath,
                                .Arguments = fullCommand,
                                .UseShellExecute = False,
                                .RedirectStandardInput = True,
                                .RedirectStandardOutput = True,
                                .RedirectStandardError = True,
                                .CreateNoWindow = True,
                                .StandardOutputEncoding = Encoding.UTF8,
                                .StandardErrorEncoding = Encoding.UTF8
                            }
                        }
                                    Return Await RunProcess(process)
                                End Using
                            End Using
                        End Using

                    Finally
                        Environment.SetEnvironmentVariable(envPassVar, Nothing, EnvironmentVariableTarget.Process)
                    End Try
                Else
                    ' Execute without credentials
                    Using process As New Process With {
                .StartInfo = New ProcessStartInfo With {
                    .FileName = "cmd.exe",
                    .Arguments = $"/c ""{taskAction.Target}""",
                    .UseShellExecute = False,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True,
                    .CreateNoWindow = True,
                    .StandardOutputEncoding = Encoding.UTF8,
                    .StandardErrorEncoding = Encoding.UTF8
                }
            }
                        Return Await RunProcess(process)
                    End Using
                End If

            Catch ex As Exception
                _logger.LogError($"Error in ExecuteBatchTask: {ex.Message}")
                Return False
            End Try
        End Function

        Public Async Function ExecuteExecutableTask(taskAction As TaskAction, credential As CredentialInfo) As Task(Of Boolean)
            Try
                _logger.LogInfo($"Executing Executable task: {taskAction.Name}")

                If credential IsNot Nothing Then
                    Dim command = If(String.IsNullOrEmpty(taskAction.Target),
                taskAction.Parameters,
                $"""{taskAction.Target}"" {taskAction.Parameters}")

                    ' Build base arguments including domain handling
                    Dim baseArgs = "-accepteula -nobanner -i"
                    If taskAction.RequiresElevation Then baseArgs &= " -h"

                    If credential.Username.Contains("\") Then
                        Dim parts = credential.Username.Split("\"c)
                        baseArgs &= $" -u {parts(0)}\{parts(1)}"
                    Else
                        baseArgs &= $" -u {credential.Username}"
                    End If

                    _logger.LogInfo($"Using credentials for user: {credential.Username}")

                    ' Execute using environment variable for password
                    Dim envPassVar = $"EXEC_PASS_{Guid.NewGuid().ToString("N")}"
                    Try
                        Using securePass As New SecureString()
                            For Each c In New NetworkCredential("", credential.SecurePassword).Password
                                securePass.AppendChar(c)
                            Next
                            securePass.MakeReadOnly()
                            Environment.SetEnvironmentVariable(envPassVar,
                        New NetworkCredential("", securePass).Password,
                        EnvironmentVariableTarget.Process)
                        End Using

                        Dim fullCommand = $"{baseArgs} -p ""%{envPassVar}%"" {command}"
                        Using process As New Process()
                            process.StartInfo = New ProcessStartInfo With {
                        .FileName = If(String.IsNullOrEmpty(taskAction.Target),
                                     "cmd.exe",
                                     taskAction.Target),
                        .Arguments = fullCommand,
                        .UseShellExecute = False,
                        .RedirectStandardOutput = True,
                        .RedirectStandardError = True,
                        .CreateNoWindow = True,
                        .StandardOutputEncoding = Encoding.GetEncoding(863),
                        .StandardErrorEncoding = Encoding.GetEncoding(863)
                    }

                            Return Await RunProcess(process)
                        End Using

                    Finally
                        Environment.SetEnvironmentVariable(envPassVar, Nothing, EnvironmentVariableTarget.Process)
                    End Try
                Else
                    ' Execute without credentials
                    Using process As New Process With {
                .StartInfo = New ProcessStartInfo With {
                    .FileName = If(String.IsNullOrEmpty(taskAction.Target),
                                 "cmd.exe",
                                 taskAction.Target),
                    .Arguments = If(String.IsNullOrEmpty(taskAction.Target),
                                  $"/c {taskAction.Parameters}",
                                  taskAction.Parameters),
                    .UseShellExecute = False,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True,
                    .CreateNoWindow = True,
                    .StandardOutputEncoding = Encoding.GetEncoding(863),
                    .StandardErrorEncoding = Encoding.GetEncoding(863)
                }
            }
                        Return Await RunProcess(process)
                    End Using
                End If

            Catch ex As Exception
                _logger.LogError($"Error in ExecuteExecutableTask: {ex.Message}")
                Return False
            End Try
        End Function

        Private Function ExtractSqlInfo(sqlContent As String) As (ServerName As String, DatabaseName As String)
            Try
                ' Look for SQL file comments first
                Dim serverMatch = Regex.Match(sqlContent, "(?i)(?:--\s*|/\*\s*)Server\s*[:=]\s*([^*\r\n]+?)(?:\s*\*/|\r|\n)")
                Dim dbMatch = Regex.Match(sqlContent, "(?i)(?:--\s*|/\*\s*)Database\s*[:=]\s*([^*\r\n]+?)(?:\s*\*/|\r|\n)")

                Dim server = If(serverMatch.Success, serverMatch.Groups(1).Value.Trim(), "")
                Dim database = If(dbMatch.Success, dbMatch.Groups(1).Value.Trim(), "")

                ' Check for USE statement if database not found in comments
                If String.IsNullOrEmpty(database) Then
                    Dim useMatch = Regex.Match(sqlContent, "USE\s+\[?([^\]\s;]+)\]?\s*;?", RegexOptions.IgnoreCase)
                    If useMatch.Success Then
                        database = useMatch.Groups(1).Value
                    End If
                End If

                _logger.LogInfo($"Extracted from SQL file - Server: {If(String.IsNullOrEmpty(server), "Not found", server)}, Database: {If(String.IsNullOrEmpty(database), "Not found", database)}")
                Return (server, database)
            Catch ex As Exception
                _logger.LogInfo($"No information was found in the SQL file or an error occurred: {ex.Message}")
                Return ("", "")
            End Try
        End Function

        Private Async Function ExecuteProcessWithOutput(process As Process, Optional verboseMode As Boolean = False) As Task(Of String)
            Dim output As New StringBuilder()
            Dim err As New StringBuilder()

            process.StartInfo.StandardOutputEncoding = Encoding.UTF8
            process.StartInfo.StandardErrorEncoding = Encoding.UTF8

            AddHandler process.OutputDataReceived, Sub(sender, e)
                                                       If e.Data IsNot Nothing Then
                                                           output.AppendLine(e.Data)
                                                           If verboseMode Then
                                                               _logger.LogInfo($"Process output: {e.Data}")
                                                           End If
                                                       End If
                                                   End Sub

            AddHandler process.ErrorDataReceived, Sub(sender, e)
                                                      If e.Data IsNot Nothing Then
                                                          If e.Data.Contains("Connecting") OrElse e.Data.Contains("Starting") OrElse
               e.Data.Contains("Copying") OrElse e.Data.Contains("exited") Then
                                                              _logger.LogInfo($"Process status: {e.Data}")
                                                          ElseIf Not String.IsNullOrWhiteSpace(e.Data) Then
                                                              err.AppendLine(e.Data)
                                                              If verboseMode Then
                                                                  _logger.LogWarning($"Process error: {e.Data}")
                                                              End If
                                                          End If
                                                      End If
                                                  End Sub

            process.Start()
            process.BeginOutputReadLine()
            process.BeginErrorReadLine()
            Await process.WaitForExitAsync()

            If process.ExitCode <> 0 Then
                Dim errorMessage = err.ToString()
                Return If(String.IsNullOrWhiteSpace(errorMessage),
            $"Error: Process exited with code {process.ExitCode}",
            $"Error: {errorMessage}")
            End If

            Return output.ToString()
        End Function

        Private Async Function ExecuteSecureProcess(command As String, credential As CredentialInfo,
    requiresElevation As Boolean, Optional verboseMode As Boolean = True) As Task(Of ProcessResult)

            Dim envPassVar = $"PSEXEC_PASS_{Guid.NewGuid().ToString("N")}"
            Try
                ' Load PsExec from embedded resource
                Dim psExecBytes = GetEmbeddedPsExec()

                ' Set up credentials if provided
                If credential IsNot Nothing Then
                    Environment.SetEnvironmentVariable(envPassVar,
                New NetworkCredential("", credential.SecurePassword).Password,
                EnvironmentVariableTarget.Process)
                End If

                ' Build base arguments
                Dim baseArgs = "-accepteula -nobanner"
                If requiresElevation Then baseArgs &= " -h"

                If credential IsNot Nothing Then
                    If credential.Username.Contains("\") Then
                        Dim parts = credential.Username.Split("\"c)
                        baseArgs &= $" -u {parts(0)}\{parts(1)}"
                    Else
                        baseArgs &= $" -u {credential.Username}"
                    End If
                    baseArgs &= $" -p ""%{envPassVar}%"""
                End If

                baseArgs &= $" {command}"

                ' Execute PsExec from memory without writing to disk
                Using ms As New MemoryStream(psExecBytes)
                    Using process As New Process()
                        process.StartInfo = New ProcessStartInfo With {
                    .FileName = "rundll32.exe",
                    .Arguments = $"kernel32.dll,LoadLibrary ""{ms.ToString()}"" {baseArgs}",
                    .UseShellExecute = False,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True,
                    .CreateNoWindow = True,
                    .StandardOutputEncoding = Encoding.UTF8,
                    .StandardErrorEncoding = Encoding.UTF8
                }

                        Dim output As New StringBuilder()
                        Dim err As New StringBuilder()

                        AddHandler process.OutputDataReceived, Sub(sender, e)
                                                                   If e.Data IsNot Nothing Then
                                                                       output.AppendLine(e.Data)
                                                                       If verboseMode Then
                                                                           _logger.LogInfo($"Process output: {e.Data}")
                                                                       End If
                                                                   End If
                                                               End Sub

                        AddHandler process.ErrorDataReceived, Sub(sender, e)
                                                                  If e.Data IsNot Nothing Then
                                                                      If e.Data.Contains("Connecting") OrElse
                           e.Data.Contains("Starting") OrElse
                           e.Data.Contains("Copying") OrElse
                           e.Data.Contains("exited") Then
                                                                          _logger.LogInfo($"Process status: {e.Data}")
                                                                      ElseIf Not String.IsNullOrWhiteSpace(e.Data) Then
                                                                          err.AppendLine(e.Data)
                                                                          If verboseMode Then
                                                                              _logger.LogWarning($"Process error: {e.Data}")
                                                                          End If
                                                                      End If
                                                                  End If
                                                              End Sub

                        process.Start()
                        process.BeginOutputReadLine()
                        process.BeginErrorReadLine()
                        Await process.WaitForExitAsync()

                        Return New ProcessResult With {
                    .ExitCode = process.ExitCode,
                    .Output = output.ToString(),
                    .GotError = err.ToString()
                }
                    End Using
                End Using

            Finally
                If credential IsNot Nothing Then
                    Environment.SetEnvironmentVariable(envPassVar, Nothing, EnvironmentVariableTarget.Process)
                End If
            End Try
        End Function

        Public Async Function ExecuteSqlCommandWithSqlCmd(query As String, server As String, database As String, credential As CredentialInfo) As Task(Of String)
            Try
                _logger.LogInfo($"Executing SQL command on {server}.{database} using sqlcmd")

                Dim tempDir = Path.Combine(Application.StartupPath, "LiteTaskData", "temp")
                Directory.CreateDirectory(tempDir)
                Dim tempFile = Path.Combine(tempDir, $"sqlcmd_{DateTime.Now:yyyyMMddHHmmss}.sql")

                Try
                    Await File.WriteAllTextAsync(tempFile, query)
                    _logger.LogInfo($"Created temporary SQL file: {tempFile}")

                    Dim sqlcmdPath = Path.Combine(_toolManager._toolsPath, "sqlcmd.exe")
                    Dim sqlcmdArgs As String

                    If credential IsNot Nothing Then
                        sqlcmdArgs = $"""-S"" ""{server}"" ""-d"" ""{database}"" ""-I"" ""-h-1"" ""-i"" ""{tempFile}"" -U ""{credential.Username}"" -P ""{credential.Password}"""
                        _logger.LogInfo($"Executing with SQL credentials for user: {credential.Username}")
                    Else
                        sqlcmdArgs = $"""-S"" ""{server}"" ""-d"" ""{database}"" ""-I"" ""-h-1"" ""-i"" ""{tempFile}"" -E"
                        _logger.LogInfo("Executing with integrated security")
                    End If

                    Using process As New Process With {
                    .StartInfo = New ProcessStartInfo(sqlcmdPath) With {
                        .Arguments = sqlcmdArgs,
                        .UseShellExecute = False,
                        .RedirectStandardInput = True,
                        .RedirectStandardOutput = True,
                        .RedirectStandardError = True,
                        .CreateNoWindow = True,
                        .WorkingDirectory = _toolManager._toolsPath,
                        .StandardOutputEncoding = Encoding.UTF8,
                        .StandardErrorEncoding = Encoding.UTF8
                    }
                }
                            Return Await ExecuteProcessWithOutput(process)
                        End Using

                Finally
                    If File.Exists(tempFile) Then
                        Try
                            File.Delete(tempFile)
                            _logger.LogInfo($"Deleted temporary SQL file: {tempFile}")
                        Catch ex As Exception
                            _logger.LogWarning($"Failed to delete temp file {tempFile}: {ex.Message}")
                        End Try
                    End If
                End Try

            Catch ex As Exception
            _logger.LogError($"Error executing SQL command: {ex.Message}")
            Throw
            End Try
        End Function

        Async Function ExecuteSqlTask(taskAction As TaskAction, credential As CredentialInfo) As Task(Of Boolean)
            Try
                _logger.LogInfo($"Executing SQL task: {taskAction.Name}")
                Dim sqlContent = File.ReadAllText(taskAction.Target)
                Dim sqlInfo = ExtractSqlInfo(sqlContent)
                Dim sqlConfig = _xmlManager.GetSqlConfiguration()
                Dim parameters = If(Not String.IsNullOrEmpty(taskAction.Parameters),
                          ParseSqlParameters(taskAction.Parameters),
                          New Dictionary(Of String, String))

                Dim server = If(parameters.ContainsKey("server"), parameters("server"),
                    If(Not String.IsNullOrEmpty(sqlInfo.ServerName), sqlInfo.ServerName,
                    sqlConfig("DefaultServer")))

                Dim database = If(parameters.ContainsKey("database"), parameters("database"),
                       If(Not String.IsNullOrEmpty(sqlInfo.DatabaseName), sqlInfo.DatabaseName,
                       sqlConfig("DefaultDatabase")))

                If String.IsNullOrEmpty(server) OrElse String.IsNullOrEmpty(database) Then
                    Throw New ArgumentException("Server and database must be specified in either task parameters, SQL file, or in configuration")
                End If

                _logger.LogInfo($"Using Server: {server}, Database: {database}")
                _logger.LogInfo($"Using credentials: {If(credential IsNot Nothing, credential.Username, "None")}")

                ' Execute using sqlcmd
                Dim result = Await ExecuteSqlCommandWithSqlCmd(sqlContent, server, database, credential)
                Return True
            Catch ex As Exception
                _logger.LogError($"Error in ExecuteSqlTask: {ex.Message}")
                Return False
            End Try
        End Function

        Public Async Function ExecuteStoredProcedureWithSqlCmd(
            storedProcName As String,
            server As String,
            database As String,
            parameters As Dictionary(Of String, Object),
            credential As CredentialInfo) As Task(Of String)

            Dim tempSqlFile As String = Nothing
            Try
                ' Create temp directory
                Dim tempDir = Path.Combine(Application.StartupPath, "LiteTaskData", "temp")
                Directory.CreateDirectory(tempDir)
                tempSqlFile = Path.Combine(tempDir, $"sp_{DateTime.Now:yyyyMMddHHmmss}.sql")

                ' Build SQL script
                Dim sqlScript As New StringBuilder()
                sqlScript.AppendLine($"USE [{database}]")
                sqlScript.AppendLine("GO")
                sqlScript.AppendLine("SET NOCOUNT ON")
                sqlScript.AppendLine("GO")

                ' Add parameters if provided
                If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
                    For Each param In parameters
                        sqlScript.AppendLine($"DECLARE {param.Key} = {param.Value}")
                    Next
                End If

                sqlScript.AppendLine("DECLARE @return_value int")
                sqlScript.AppendLine($"EXEC @return_value = {storedProcName}")
                sqlScript.AppendLine("SELECT 'Return Value' = @return_value")
                sqlScript.AppendLine("GO")

                ' Write SQL file
                Await File.WriteAllTextAsync(tempSqlFile, sqlScript.ToString())

                ' Execute using sqlcmd
                Return Await ExecuteSqlCommandWithSqlCmd(
            File.ReadAllText(tempSqlFile),
            server,
            database,
            credential)

            Finally
                If tempSqlFile IsNot Nothing AndAlso File.Exists(tempSqlFile) Then
                    Try
                        File.Delete(tempSqlFile)
                    Catch ex As Exception
                        _logger.LogWarning($"Failed to delete temp file {tempSqlFile}: {ex.Message}")
                    End Try
                End If
            End Try
        End Function

        Public Async Function ExecutePowerShellTask(taskAction As TaskAction, credential As CredentialInfo) As Task(Of Boolean)
            Try
                _logger.LogInfo($"Starting PowerShell task execution: {taskAction.Name}")
                _logger.LogInfo($"Script path: {taskAction.Target}")
                _logger.LogInfo($"Using credentials: {If(credential IsNot Nothing, credential.Username, "None")}")

                Dim scriptContent = File.ReadAllText(taskAction.Target)

                ' Detect and verify required modules before execution
                Dim requiredModules = PowerShellPathManager.DetectRequiredModules(scriptContent)
                If requiredModules.Count > 0 Then
                    _logger.LogInfo($"Detected required PowerShell modules: {String.Join(", ", requiredModules)}")
                    Dim missingModules = _powerShellPathManager.VerifyModules(requiredModules)
                    If missingModules.Count > 0 Then
                        _logger.LogWarning($"Missing PowerShell modules: {String.Join(", ", missingModules)}. Script may fail if these modules cannot be auto-loaded. Install them via the Modules Manager or run: Install-Module {String.Join(", ", missingModules)}")
                    End If
                End If

                Using powerShell As PowerShell = _powerShellPathManager.CreatePowerShellInstance()
                    powerShell.AddScript(scriptContent)

                    If credential IsNot Nothing Then
                        _logger.LogInfo("Adding credential parameters to PowerShell script")
                        powerShell.AddParameter("username", credential.Username)
                        ' Create new SecureString from password to avoid disposed object
                        Dim securePass = New NetworkCredential("", credential.Password).SecurePassword
                        powerShell.AddParameter("password", New NetworkCredential("", securePass).Password)
                    End If

                    If Not String.IsNullOrEmpty(taskAction.Parameters) Then
                        _logger.LogInfo($"Adding parameters: {taskAction.Parameters}")
                        For Each param In ParseParameters(taskAction.Parameters)
                            powerShell.AddParameter(param.Key, param.Value)
                        Next
                    End If

                    Dim result = Await powerShell.InvokeAsync()
                    _logger.LogInfo($"PowerShell execution completed. Output count: {result?.Count}")

                    ' Log all output and errors
                    For Each item In result
                        _logger.LogInfo($"PowerShell output: {item}")
                    Next

                    For Each info In powerShell.Streams.Information
                        _logger.LogInfo($"PowerShell information: {info}")
                    Next

                    If powerShell.Streams.Error.Count > 0 Then
                        For Each err As ErrorRecord In powerShell.Streams.Error
                            _logger.LogError($"PowerShell error: {err.Exception.Message}")
                            _logger.LogError($"PowerShell error details: {err.ScriptStackTrace}")
                        Next
                        Return False
                    End If

                    Return True  ' If no errors, consider it successful
                End Using

            Catch ex As Exception
                _logger.LogError($"Error in ExecutePowerShellTask: {ex.Message}")
                _logger.LogError($"Stack trace: {ex.StackTrace}")
                Return False
            End Try
        End Function



        Private Function ParseParameters(parameters As String) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)
            Try
                If String.IsNullOrEmpty(parameters) Then Return result

                ' Use Span for more efficient string splitting (if available) or stick with optimized approach
                Dim _paramArray = parameters.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                For Each param In _paramArray
                    Dim equalIndex = param.IndexOf("="c)
                    If equalIndex > 0 AndAlso equalIndex < param.Length - 1 Then
                        Dim key = param.Substring(0, equalIndex)
                        Dim value = param.Substring(equalIndex + 1)
                        result(key) = value
                    End If
                Next

                _logger.LogInfo($"Parsed {result.Count} parameters successfully")
                Return result
            Catch ex As Exception
                _logger.LogError($"Error parsing parameters: {ex.Message}")
                Return result
            End Try
        End Function

        Private Function ParseSqlParameters(command As String) As Dictionary(Of String, Object)
            Dim parameters As New Dictionary(Of String, Object)
            Try
                Dim paramMatch = Regex.Match(command, "EXEC(?:UTE)?\s+\w+\s*(.+)", RegexOptions.IgnoreCase)
                If paramMatch.Success Then
                    For Each param In paramMatch.Groups(1).Value.Split(","c)
                        Dim parts = param.Trim().Split("="c)
                        If parts.Length = 2 Then
                            Dim key = parts(0).Trim().TrimStart("@"c)
                            Dim value = parts(1).Trim().Trim("'"c)
                            parameters.Add(key, value)
                        End If
                    Next
                End If
            Catch ex As Exception
                _logger?.LogError($"Error parsing SQL parameters: {ex.Message}")
            End Try
            Return parameters
        End Function

        Private Function ParseUNCPaths(batchContent As String) As List(Of String)
            Dim paths As New List(Of String)

            ' Match UNC paths: \\server\share
            Dim regex = New Regex("\\\\[^\\]+\\[^\\]+")
            Dim matches = regex.Matches(batchContent)

            For Each match As Match In matches
                paths.Add(match.Value)
            Next

            Return paths.Distinct().ToList()
        End Function

        Private Function GetEmbeddedPsExec() As Byte()
            Using stream As Stream = GetType(TaskRunner).Assembly.GetManifestResourceStream("LiteTask.PsExec64.exe")
                If stream Is Nothing Then
                    Throw New InvalidOperationException("Embedded PsExec resource not found")
                End If
                Dim buffer As Byte() = New Byte(stream.Length - 1) {}
                stream.Read(buffer, 0, buffer.Length)
                Return buffer
            End Using
        End Function

        Private Function MaskSensitiveData(input As String, password As String) As String
            If String.IsNullOrEmpty(input) OrElse String.IsNullOrEmpty(password) Then Return input
            Return input.Replace(password, "[MASKED]")
        End Function

        Private Sub OnErrorReceived(sender As Object, e As DataAddedEventArgs)
            Dim errorRecord As ErrorRecord = CType(sender, PSDataCollection(Of ErrorRecord))(e.Index)
            _logger.LogError($"PowerShell Error: {errorRecord.Exception.Message}")
        End Sub

        Private Sub OnInformationReceived(sender As Object, e As DataAddedEventArgs)
            Dim informationRecord As InformationRecord = CType(sender, PSDataCollection(Of InformationRecord))(e.Index)
            _logger.LogInfo($"PowerShell Output: {informationRecord.MessageData}")
        End Sub

        Private Async Function RunProcess(process As Process) As Task(Of Boolean)
            Dim output As New StringBuilder()
            Dim errors As New StringBuilder()

            Try
                AddHandler process.OutputDataReceived, Sub(sender, e)
                                                           If e.Data IsNot Nothing Then
                                                               output.AppendLine(e.Data)
                                                               _logger.LogInfo($"Process output: {e.Data}")
                                                           End If
                                                       End Sub

                AddHandler process.ErrorDataReceived, Sub(sender, e)
                                                          If e.Data IsNot Nothing Then
                                                              errors.AppendLine(e.Data)
                                                              _logger.LogWarning($"Process error: {e.Data}")
                                                          End If
                                                      End Sub

                process.Start()
                process.BeginOutputReadLine()
                process.BeginErrorReadLine()
                Await process.WaitForExitAsync()

                If process.ExitCode <> 0 Then
                    _logger.LogError($"Process failed with exit code: {process.ExitCode}")
                    _logger.LogError($"Error output: {errors.ToString()}")
                    Return False
                End If

                Return True

            Catch ex As Exception
                _logger.LogError($"Error running process: {ex.Message}")
                Return False
            End Try
        End Function

        Private Sub VerifyRequiredTools()
            ' Check for required tools
            Dim requiredTools = New String() {"sqlcmd.exe"}
            Dim missingTools = New List(Of String)

            For Each tool In requiredTools
                Dim toolPath = Path.Combine(_toolManager._toolsPath, tool)
                If Not File.Exists(toolPath) Then
                    missingTools.Add(tool)
                End If
            Next

            If missingTools.Any() Then
                Throw New Exception($"Required tools missing: {String.Join(", ", missingTools)}. Please ensure all required tools are in the tools directory: {_toolManager._toolsPath}")
            End If
        End Sub

        ''' Cleans up old temporary files from previous runs
        Public Sub CleanupOldTempFiles()
            Try
                Dim tempDir = Path.Combine(Application.StartupPath, "LiteTaskData", "temp")
                If Directory.Exists(tempDir) Then
                    Dim cutoffTime = DateTime.Now.AddHours(-1) ' Delete files older than 1 hour
                    
                    ' Clean up SQL temp files
                    Dim sqlTempFiles = Directory.GetFiles(tempDir, "*.sql")
                    For Each _file In sqlTempFiles
                        Try
                            Dim fileInfo = New FileInfo(_file)
                            If fileInfo.LastWriteTime < cutoffTime Then
                                File.Delete(_file)
                                _logger.LogInfo($"Cleaned up old SQL temp file: {_file}")
                            End If
                        Catch ex As Exception
                            _logger.LogWarning($"Failed to delete old SQL temp file {_file}: {ex.Message}")
                        End Try
                    Next
                    
                    ' Clean up XML temp files
                    Dim xmlTempFiles = Directory.GetFiles(tempDir, "xml_save_*.tmp")
                    For Each _file In xmlTempFiles
                        Try
                            Dim fileInfo = New FileInfo(_file)
                            If fileInfo.LastWriteTime < cutoffTime Then
                                File.Delete(_file)
                                _logger.LogInfo($"Cleaned up old XML temp file: {_file}")
                            End If
                        Catch ex As Exception
                            _logger.LogWarning($"Failed to delete old XML temp file {_file}: {ex.Message}")
                        End Try
                    Next
                    
                    ' Clean up log rotation temp files
                    Dim logTempFiles = Directory.GetFiles(tempDir, "log_rotation_*.tmp")
                    For Each _file In logTempFiles
                        Try
                            Dim fileInfo = New FileInfo(_file)
                            If fileInfo.LastWriteTime < cutoffTime Then
                                File.Delete(_file)
                                _logger.LogInfo($"Cleaned up old log rotation temp file: {_file}")
                            End If
                        Catch ex As Exception
                            _logger.LogWarning($"Failed to delete old log rotation temp file {_file}: {ex.Message}")
                        End Try
                    Next
                End If
            Catch ex As Exception
                _logger.LogError($"Error during temp file cleanup: {ex.Message}")
            End Try
        End Sub

    End Class

    Public Class SqlCommandSanitizer
        Public Function SanitizeColumnName(columnName As String) As String
            If String.IsNullOrEmpty(columnName) Then Return String.Empty
            Return Regex.Replace(columnName, "[^\w\s\-\[\]]", "")
        End Function

        Public Function SanitizeOutput(output As String) As String
            If String.IsNullOrEmpty(output) Then Return String.Empty
            Return Regex.Replace(output, "[\x00-\x08\x0B\x0C\x0E-\x1F]", "")
        End Function
    End Class

    Public Class SqlExecutionResult
        Public Property Success As Boolean
        Public Property Message As String
        Public Property Data As DataTable
        Public Property RowsAffected As Integer
    End Class

    Public Class TempFile
        Implements IDisposable

        Public ReadOnly Property Path As String

        Public Sub New(extension As String)
            ' Create temp directory if it doesn't exist
            Dim tempDir = System.IO.Path.Combine(Application.StartupPath, "LiteTaskData", "temp")
            Directory.CreateDirectory(tempDir)

            ' Generate unique temp file path in our custom directory
            Path = System.IO.Path.Combine(tempDir, Guid.NewGuid().ToString() & extension)
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If File.Exists(Path) Then
                Try
                    File.Delete(Path)
                Catch
                    ' Log but don't throw
                End Try
            End If
        End Sub
    End Class

End Namespace
