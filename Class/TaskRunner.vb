Imports System.Collections.Concurrent
Imports System.Collections.ObjectModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel

Namespace LiteTask
    Public Class TaskRunner
        Implements IDisposable

        ' TODO: Phase 2 - Add audit logging for database operations

        Private ReadOnly _logger As Logger
        Private ReadOnly _credentialManager As CredentialManager
        Private ReadOnly _liteRunConfig As LiteRunConfig
        Private ReadOnly _toolManager As ToolManager
        Private ReadOnly _xmlManager As XMLManager
        Private ReadOnly _logPath As String
        Private ReadOnly _powerShellPathManager As PowerShellPathManager
        Private ReadOnly _sqlTab As SqlTab
        Private ReadOnly _sanitizer As New SqlCommandSanitizer()
        Private _connectionStringBase As String
        Private _sqlConfig As Dictionary(Of String, String)
        Private _disposed As Boolean
        Public Event OutputReceived(sender As Object, data As String)
        Public Event ErrorReceived(sender As Object, data As String)

        ' Live child processes per execution context (task name), so the scheduler
        ' can locate and kill the owning process when a task hangs past its timeout
        ' or its lock goes stale. AsyncLocal flows the task name from the scheduler
        ' into the nested async execution path without changing method signatures.
        Private ReadOnly _trackedProcesses As New ConcurrentDictionary(Of String, ConcurrentDictionary(Of Integer, Process))
        Private ReadOnly _currentTaskName As New Threading.AsyncLocal(Of String)
        Private Const ADHOC_CONTEXT As String = "_adhoc"
        Private Const DEFAULT_TASK_TIMEOUT_MINUTES As Integer = 60

        ''' <summary>
        ''' Execution context for process tracking. Set by the scheduler before
        ''' running a task's actions; flows through awaits into RunProcessWithTimeout.
        ''' </summary>
        Public Property CurrentTaskName As String
            Get
                Return _currentTaskName.Value
            End Get
            Set(value As String)
                _currentTaskName.Value = value
            End Set
        End Property

        ''' <summary>
        ''' Effective hard timeout for a task action. TimeoutMinutes has always been
        ''' persisted per action but was never enforced; values &lt;= 0 fall back to
        ''' the default so every run has a bounded duration.
        ''' </summary>
        Public Shared Function GetActionTimeout(taskAction As TaskAction) As TimeSpan
            Dim minutes = If(taskAction IsNot Nothing AndAlso taskAction.TimeoutMinutes > 0,
                             taskAction.TimeoutMinutes,
                             DEFAULT_TASK_TIMEOUT_MINUTES)
            Return TimeSpan.FromMinutes(minutes)
        End Function

        Private Function TrackProcess(process As Process) As String
            Try
                Dim key = If(CurrentTaskName, ADHOC_CONTEXT)
                Dim bucket = _trackedProcesses.GetOrAdd(key, Function(k) New ConcurrentDictionary(Of Integer, Process)())
                bucket(process.Id) = process
                _logger.LogInfo($"Tracking child process {process.Id} (context: {key})")
                Return key
            Catch ex As Exception
                ' Process may have exited before Id could be read; nothing to track
                _logger.LogWarning($"Could not track child process: {ex.Message}")
                Return Nothing
            End Try
        End Function

        Private Sub UntrackProcess(key As String, process As Process)
            If key Is Nothing Then Return
            Try
                Dim bucket As ConcurrentDictionary(Of Integer, Process) = Nothing
                If _trackedProcesses.TryGetValue(key, bucket) Then
                    bucket.TryRemove(process.Id, Nothing)
                End If
            Catch
                ' Id unavailable (process never started) - nothing was tracked
            End Try
        End Sub

        ''' <summary>
        ''' Kills every tracked child process (entire process tree) started for the
        ''' given task. Used by the scheduler's timeout and stale-lock recovery so a
        ''' hung executable cannot hold the task lock indefinitely.
        ''' </summary>
        Public Function KillProcessesForTask(taskName As String) As Integer
            Dim killed = 0
            Dim bucket As ConcurrentDictionary(Of Integer, Process) = Nothing
            If Not _trackedProcesses.TryGetValue(taskName, bucket) OrElse bucket.IsEmpty Then
                _logger.LogWarning($"No tracked child process found for task '{taskName}'")
                Return 0
            End If

            For Each kvp In bucket
                Try
                    If Not kvp.Value.HasExited Then
                        _logger.LogWarning($"Killing hung process {kvp.Key} (and children) for task '{taskName}'")
                        kvp.Value.Kill(entireProcessTree:=True)
                        killed += 1
                    End If
                Catch ex As ObjectDisposedException
                    ' The Process wrapper was disposed but the OS process may live on: kill by PID
                    Try
                        Using p = Process.GetProcessById(kvp.Key)
                            _logger.LogWarning($"Killing hung process {kvp.Key} (by PID) for task '{taskName}'")
                            p.Kill(entireProcessTree:=True)
                            killed += 1
                        End Using
                    Catch
                        ' Process already gone
                    End Try
                Catch ex As Exception
                    _logger.LogError($"Failed to kill process {kvp.Key} for task '{taskName}': {ex.Message}")
                Finally
                    bucket.TryRemove(kvp.Key, Nothing)
                End Try
            Next

            Return killed
        End Function

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
            _sqlConfig = _xmlManager.GetSqlConfiguration()

            ' Ensure required directories exist
            Directory.CreateDirectory(_logPath)
            Directory.CreateDirectory(_toolManager._toolsPath)

            ' Verify required tools
            VerifyRequiredTools()

            ' Initialize PowerShell environment
            _powerShellPathManager.EnsureModulePathExists()

            ' Make sure PsExec64.exe is available in the tools folder (best effort:
            ' a failure here is retried on first credentialed Batch/Executable run)
            Try
                GetPsExecPath()
            Catch ex As Exception
                _logger.LogWarning($"Could not extract embedded PsExec64.exe: {ex.Message}")
            End Try

            ' Clean up any old temporary files from previous runs
            CleanupOldTempFiles()

        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If Not _disposed Then
                _sqlConfig = Nothing
                _disposed = True
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

                ' Trailing args appended to `cmd.exe /c "target.bat" <args>` so the
                ' script receives them as %1..%* (works for both code paths).
                Dim batchArgs As String = If(String.IsNullOrWhiteSpace(taskAction.Parameters),
                                             String.Empty,
                                             " " & taskAction.Parameters.Trim())

                If credential IsNot Nothing Then
                    ' Parse UNC paths from batch content
                    Dim batchContent = File.ReadAllText(taskAction.Target)
                    Dim uncPaths = ParseUNCPaths(batchContent)

                    ' Write the augmented script (net use prelude + original body) to a
                    ' temp .bat so we can invoke it as `cmd.exe /c "tempfile" <args>`.
                    ' The previous implementation streamed the script via stdin to
                    ' `cmd /c -`, which cannot receive %1..%* — task parameters were
                    ' silently dropped.
                    Dim tempDir = Path.Combine(Application.StartupPath, "LiteTaskData", "temp")
                    Directory.CreateDirectory(tempDir)
                    Dim tempBatFile = Path.Combine(tempDir, $"batch_{Guid.NewGuid():N}.bat")

                    Try
                        ' Pre-connect the UNC paths referenced by the script. The script runs
                        ' inside the PsExec logon session of the selected account, so `net use`
                        ' authenticates with that session's own credentials: the password is
                        ' never written to the temp file. (The previous %BATCH_PASS_xxx%
                        ' placeholder could not expand there: PsExec children do not inherit
                        ' the LiteTask environment.)
                        Dim modifiedContent = New StringBuilder("@echo off" & Environment.NewLine)
                        For Each uncPath In uncPaths.Distinct()
                            modifiedContent.AppendLine($"net use {uncPath}")
                        Next
                        modifiedContent.AppendLine(batchContent)
                        For Each uncPath In uncPaths.Distinct()
                            modifiedContent.AppendLine($"net use {uncPath} /delete")
                        Next

                        Await File.WriteAllTextAsync(tempBatFile, modifiedContent.ToString())

                        Dim command = PsExecCommandBuilder.BuildBatchCommand(tempBatFile, batchArgs)
                        Return Await RunThroughPsExec(taskAction, credential, command, Encoding.UTF8)

                    Finally
                        If File.Exists(tempBatFile) Then
                            Try
                                File.Delete(tempBatFile)
                            Catch ex As Exception
                                _logger.LogWarning($"Failed to delete temp batch file {tempBatFile}: {ex.Message}")
                            End Try
                        End If
                    End Try
                Else
                    ' Execute without credentials
                    Using process As New Process With {
                .StartInfo = New ProcessStartInfo With {
                    .FileName = "cmd.exe",
                    .Arguments = $"/c ""{taskAction.Target}""{batchArgs}",
                    .UseShellExecute = False,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True,
                    .CreateNoWindow = True,
                    .StandardOutputEncoding = Encoding.UTF8,
                    .StandardErrorEncoding = Encoding.UTF8
                }
            }
                        Return Await RunProcessWithTimeout(process, GetActionTimeout(taskAction))
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
                    ' Run the target through PsExec under the selected account. The
                    ' previous implementation started the target itself (as the current
                    ' user) with PsExec arguments, so the credential was never applied.
                    Dim command = PsExecCommandBuilder.BuildExecutableCommand(taskAction.Target, taskAction.Parameters)
                    Return Await RunThroughPsExec(taskAction, credential, command, Encoding.GetEncoding(863))
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
                        Return Await RunProcessWithTimeout(process, GetActionTimeout(taskAction))
                    End Using
                End If

            Catch ex As Exception
                _logger.LogError($"Error in ExecuteExecutableTask: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Launches <paramref name="command"/> through PsExec64 as the account held in
        ''' <paramref name="credential"/>, with the action timeout and process tracking of
        ''' <see cref="RunProcessWithTimeout"/>. The password is handed to PsExec with
        ''' <c>-p</c> (PsExec offers no environment or stdin alternative); it is never
        ''' logged, and any echo of it in the child's output is masked.
        ''' </summary>
        Private Async Function RunThroughPsExec(taskAction As TaskAction,
                                                credential As CredentialInfo,
                                                command As String,
                                                outputEncoding As Encoding) As Task(Of Boolean)
            Dim psExecPath = GetPsExecPath()
            Dim password = GetPlainPassword(credential)

            _logger.LogInfo($"Using credentials for user: {credential.Username}")

            Dim fullCommand = PsExecCommandBuilder.BuildArguments(credential.Username,
                                                                 password,
                                                                 taskAction.RequiresElevation,
                                                                 Environment.UserInteractive,
                                                                 command)

            Using process As New Process With {
                .StartInfo = New ProcessStartInfo With {
                    .FileName = psExecPath,
                    .Arguments = fullCommand,
                    .UseShellExecute = False,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True,
                    .CreateNoWindow = True,
                    .StandardOutputEncoding = outputEncoding,
                    .StandardErrorEncoding = outputEncoding
                }
            }
                Return Await RunProcessWithTimeout(process, GetActionTimeout(taskAction), sensitiveValue:=password)
            End Using
        End Function

        Private Shared Function GetPlainPassword(credential As CredentialInfo) As String
            If credential.SecurePassword IsNot Nothing AndAlso credential.SecurePassword.Length > 0 Then
                Return New NetworkCredential(String.Empty, credential.SecurePassword).Password
            End If
            Return If(credential.Password, String.Empty)
        End Function

        ''' <summary>
        ''' Full path of PsExec64.exe in the tools folder, extracting the embedded copy
        ''' on first use (nothing else writes it to disk).
        ''' </summary>
        Private Function GetPsExecPath() As String
            Dim psExecPath = Path.Combine(_toolManager._toolsPath, "PsExec64.exe")
            If Not File.Exists(psExecPath) Then
                Directory.CreateDirectory(_toolManager._toolsPath)
                File.WriteAllBytes(psExecPath, GetEmbeddedPsExec())
                _logger.LogInfo($"Extracted embedded PsExec64.exe to {psExecPath}")
            End If
            Return psExecPath
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

        Private Async Function ExecuteProcessWithOutput(process As Process, Optional verboseMode As Boolean = False, Optional timeout As TimeSpan? = Nothing) As Task(Of String)
            Dim output As New StringBuilder()
            Dim err As New StringBuilder()
            Dim trackKey As String = Nothing
            Dim effectiveTimeout = If(timeout, TimeSpan.FromMinutes(DEFAULT_TASK_TIMEOUT_MINUTES))

            process.StartInfo.StandardOutputEncoding = Encoding.UTF8
            process.StartInfo.StandardErrorEncoding = Encoding.UTF8

            ' Named handlers so they can be removed after execution (prevents closure leaks)
            Dim outputHandler As DataReceivedEventHandler = Sub(sender, e)
                                                                If e.Data IsNot Nothing Then
                                                                    output.AppendLine(e.Data)
                                                                    If verboseMode Then
                                                                        _logger.LogInfo($"Process output: {e.Data}")
                                                                    End If
                                                                End If
                                                            End Sub

            Dim errorHandler As DataReceivedEventHandler = Sub(sender, e)
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

            Try
                AddHandler process.OutputDataReceived, outputHandler
                AddHandler process.ErrorDataReceived, errorHandler

                process.Start()
                process.BeginOutputReadLine()
                process.BeginErrorReadLine()
                trackKey = TrackProcess(process)

                ' Enforce a hard timeout so a hung child (e.g. sqlcmd waiting on an
                ' unavailable SQL Server) cannot hold the task lock forever
                Using cts As New CancellationTokenSource(effectiveTimeout)
                    Try
                        Await process.WaitForExitAsync(cts.Token)
                    Catch ex As OperationCanceledException
                        _logger.LogError($"Process exceeded {effectiveTimeout.TotalMinutes:F0}-minute timeout. Killing process tree (PID {process.Id}).")
                        Try
                            process.Kill(entireProcessTree:=True)
                        Catch killEx As Exception
                            _logger.LogWarning($"Error killing timed-out process: {killEx.Message}")
                        End Try
                        Return $"Error: Process exceeded {effectiveTimeout.TotalMinutes:F0}-minute timeout and was killed"
                    End Try
                End Using

                If process.ExitCode <> 0 Then
                    Dim errorMessage = err.ToString()
                    Return If(String.IsNullOrWhiteSpace(errorMessage),
                $"Error: Process exited with code {process.ExitCode}",
                $"Error: {errorMessage}")
                End If

                Return output.ToString()
            Finally
                RemoveHandler process.OutputDataReceived, outputHandler
                RemoveHandler process.ErrorDataReceived, errorHandler

                ' Kill the process if it's still running to prevent orphaned processes
                Try
                    If Not process.HasExited Then
                        process.Kill(entireProcessTree:=True)
                    End If
                Catch
                    ' Suppress errors during cleanup
                End Try

                UntrackProcess(trackKey, process)
            End Try
        End Function

        Public Async Function ExecuteSqlCommandWithSqlCmd(query As String, server As String, database As String, credential As CredentialInfo, Optional timeout As TimeSpan? = Nothing) As Task(Of String)
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
                            Return Await ExecuteProcessWithOutput(process, timeout:=timeout)
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
                ' The task's Parameters field uses the same `key=value` / `-Name value`
                ' grammar as PowerShell tasks. ParseSqlParameters parses T-SQL EXEC
                ' syntax inside the .sql file content, not the form field.
                Dim parameters = If(Not String.IsNullOrEmpty(taskAction.Parameters),
                          ParameterParser.Parse(taskAction.Parameters),
                          New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase))

                Dim server = If(parameters.ContainsKey("server"), CStr(parameters("server")),
                    If(Not String.IsNullOrEmpty(sqlInfo.ServerName), sqlInfo.ServerName,
                    sqlConfig("DefaultServer")))

                Dim database = If(parameters.ContainsKey("database"), CStr(parameters("database")),
                       If(Not String.IsNullOrEmpty(sqlInfo.DatabaseName), sqlInfo.DatabaseName,
                       sqlConfig("DefaultDatabase")))

                If String.IsNullOrEmpty(server) OrElse String.IsNullOrEmpty(database) Then
                    Throw New ArgumentException("Server and database must be specified in either task parameters, SQL file, or in configuration")
                End If

                _logger.LogInfo($"Using Server: {server}, Database: {database}")
                _logger.LogInfo($"Using credentials: {If(credential IsNot Nothing, credential.Username, "None")}")

                ' Execute using sqlcmd, bounded by the action's configured timeout
                Dim result = Await ExecuteSqlCommandWithSqlCmd(sqlContent, server, database, credential, GetActionTimeout(taskAction))
                If result IsNot Nothing AndAlso result.StartsWith("Error:", StringComparison.OrdinalIgnoreCase) Then
                    _logger.LogError($"SQL task {taskAction.Name} failed: {result}")
                    Return False
                End If
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

                ' Create PowerShell instance with explicit Runspace tracking (PS.Dispose() doesn't reliably dispose it)
                Using powerShell As PowerShell = _powerShellPathManager.CreatePowerShellInstance()
                    Dim runspace = powerShell.Runspace

                    Try
                        ' Run init script separately, then clear so only the task script is in the pipeline
                        powerShell.Invoke()
                        powerShell.Commands.Clear()

                        powerShell.AddScript(scriptContent)

                        If credential IsNot Nothing Then
                            _logger.LogInfo("Adding credential parameters to PowerShell script")
                            powerShell.AddParameter("username", credential.Username)
                            ' Pass password directly (avoids a pointless SecureString round-trip that leaked)
                            powerShell.AddParameter("password", credential.Password)
                        End If

                        If Not String.IsNullOrEmpty(taskAction.Parameters) Then
                            _logger.LogInfo($"Adding parameters: {taskAction.Parameters}")
                            For Each param In ParseParameters(taskAction.Parameters)
                                If param.Value Is Nothing Then
                                    powerShell.AddParameter(param.Key)
                                Else
                                    powerShell.AddParameter(param.Key, param.Value)
                                End If
                            Next
                        End If

                        ' Subscribe to streams for real-time output to RunTab and other listeners
                        Dim infoHandler As EventHandler(Of DataAddedEventArgs) = Sub(sender, e)
                            Dim records = DirectCast(sender, PSDataCollection(Of InformationRecord))
                            Dim record = records(e.Index)
                            Dim message = record.MessageData?.ToString()
                            If message IsNot Nothing Then
                                _logger.LogInfo($"PowerShell information: {message}")
                                RaiseEvent OutputReceived(Me, message)
                            End If
                        End Sub

                        Dim warningHandler As EventHandler(Of DataAddedEventArgs) = Sub(sender, e)
                            Dim records = DirectCast(sender, PSDataCollection(Of WarningRecord))
                            Dim record = records(e.Index)
                            Dim message = record.Message
                            If message IsNot Nothing Then
                                _logger.LogWarning($"PowerShell warning: {message}")
                                RaiseEvent ErrorReceived(Me, $"Warning: {message}")
                            End If
                        End Sub

                        Dim errorStreamHandler As EventHandler(Of DataAddedEventArgs) = Sub(sender, e)
                            Dim records = DirectCast(sender, PSDataCollection(Of ErrorRecord))
                            Dim record = records(e.Index)
                            Dim message = record.Exception?.Message
                            If message IsNot Nothing Then
                                _logger.LogError($"PowerShell error: {message}")
                                RaiseEvent ErrorReceived(Me, message)
                            End If
                        End Sub

                        AddHandler powerShell.Streams.Information.DataAdded, infoHandler
                        AddHandler powerShell.Streams.Warning.DataAdded, warningHandler
                        AddHandler powerShell.Streams.Error.DataAdded, errorStreamHandler

                        ' Hard timeout to prevent hung scripts from holding the task lock forever
                        Dim hasErrors = False
                        Dim result As PSDataCollection(Of PSObject) = Nothing
                        Try
                            ' Hard cap from the action's configured TimeoutMinutes (default 60)
                            Dim psTimeout = GetActionTimeout(taskAction)
                            Dim invokeTask = powerShell.InvokeAsync()
                            Dim timeoutTask = Task.Delay(psTimeout)

                            If Await Task.WhenAny(invokeTask, timeoutTask) Is timeoutTask Then
                                _logger.LogError($"PowerShell task '{taskAction.Name}' exceeded {psTimeout.TotalMinutes:F0}-minute timeout. Stopping forcibly.")
                                powerShell.Stop()
                                Return False
                            End If

                            result = Await invokeTask
                            _logger.LogInfo($"PowerShell execution completed. Output count: {result?.Count}")

                            ' Log pipeline output and raise events for RunTab display
                            If result IsNot Nothing Then
                                For Each item In result
                                    Dim text = item?.ToString()
                                    If text IsNot Nothing Then
                                        _logger.LogInfo($"PowerShell output: {text}")
                                        RaiseEvent OutputReceived(Me, text)
                                    End If
                                Next
                            End If

                            ' Log any remaining error details (script stack traces)
                            If powerShell.Streams.Error.Count > 0 Then
                                For Each err As ErrorRecord In powerShell.Streams.Error
                                    _logger.LogError($"PowerShell error details: {err.ScriptStackTrace}")
                                Next
                                hasErrors = True
                            End If
                        Finally
                            ' Remove stream handlers to prevent leaking closures
                            RemoveHandler powerShell.Streams.Information.DataAdded, infoHandler
                            RemoveHandler powerShell.Streams.Warning.DataAdded, warningHandler
                            RemoveHandler powerShell.Streams.Error.DataAdded, errorStreamHandler

                            ' Dispose output collection immediately to free memory from large results
                            If result IsNot Nothing Then result.Dispose()
                        End Try

                        Return Not hasErrors

                    Finally
                        ' Unload modules to release session state (assemblies remain loaded per .NET limitation)
                        Try
                            powerShell.Commands.Clear()
                            powerShell.AddScript("Get-Module | Where-Object { $_.Name -notlike 'Microsoft.PowerShell.*' } | Remove-Module -Force -ErrorAction SilentlyContinue")
                            powerShell.Invoke()
                        Catch ex As Exception
                            _logger.LogWarning($"Error unloading PowerShell modules: {ex.Message}")
                        End Try

                        ' Clear the pipeline commands to release script references
                        Try
                            powerShell.Commands.Clear()
                        Catch ex As Exception
                            _logger.LogWarning($"Error clearing PowerShell commands: {ex.Message}")
                        End Try

                        ' Clear all streams to release object references held by the pipeline
                        Try
                            powerShell.Streams.ClearStreams()
                        Catch ex As Exception
                            _logger.LogWarning($"Error clearing PowerShell streams: {ex.Message}")
                        End Try

                        ' Explicitly close/dispose the Runspace (PS.Dispose() alone leaks it)
                        If runspace IsNot Nothing Then
                            Try
                                If runspace.RunspaceStateInfo.State <> Runspaces.RunspaceState.Closed Then
                                    runspace.Close()
                                End If
                                runspace.Dispose()
                            Catch ex As Exception
                                _logger.LogWarning($"Error disposing PowerShell runspace: {ex.Message}")
                            End Try
                        End If

                        ' Force a full GC to reclaim COM interop objects and large ImportExcel/EPPlus
                        ' object graphs that the non-blocking Optimized collect won't release in time.
                        ' WaitForPendingFinalizers ensures COM Release() calls complete before we return.
                        Try
                            GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking:=True)
                            GC.WaitForPendingFinalizers()
                            GC.Collect(2, GCCollectionMode.Forced, blocking:=False)
                        Catch
                            ' GC errors should never propagate
                        End Try
                    End Try
                End Using

            Catch ex As Exception
                _logger.LogError($"Error in ExecutePowerShellTask: {ex.Message}")
                _logger.LogError($"Stack trace: {ex.StackTrace}")
                Return False
            End Try
        End Function



        Private Function ParseParameters(parameters As String) As Dictionary(Of String, Object)
            Try
                Dim result = ParameterParser.Parse(parameters)
                _logger.LogInfo($"Parsed {result.Count} parameters successfully")
                Return result
            Catch ex As Exception
                _logger.LogError($"Error parsing parameters: {ex.Message}")
                Return New Dictionary(Of String, Object)
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

        ''' <summary>
        ''' Runs a process with a hard timeout. If the process exceeds the timeout,
        ''' it is killed (including the entire process tree) and False is returned.
        ''' The process is tracked by PID under the current task context so the
        ''' scheduler's stale-lock recovery can kill it if this method itself hangs.
        ''' When <paramref name="sensitiveValue"/> is supplied (e.g. a password passed
        ''' on the child's command line), every occurrence in the captured output is
        ''' masked before it reaches the log or the GUI.
        ''' </summary>
        Private Async Function RunProcessWithTimeout(process As Process, timeout As TimeSpan,
                                                     Optional sensitiveValue As String = Nothing) As Task(Of Boolean)
            Dim output As New StringBuilder()
            Dim errors As New StringBuilder()
            Dim trackKey As String = Nothing

            Dim outputHandler As DataReceivedEventHandler = Sub(sender, e)
                                                                If e.Data IsNot Nothing Then
                                                                    Dim line = MaskSensitiveData(e.Data, sensitiveValue)
                                                                    output.AppendLine(line)
                                                                    _logger.LogInfo($"Process output: {line}")
                                                                    RaiseEvent OutputReceived(Me, line)
                                                                End If
                                                            End Sub

            Dim errorHandler As DataReceivedEventHandler = Sub(sender, e)
                                                               If e.Data IsNot Nothing Then
                                                                   Dim line = MaskSensitiveData(e.Data, sensitiveValue)
                                                                   errors.AppendLine(line)
                                                                   _logger.LogWarning($"Process error: {line}")
                                                                   RaiseEvent ErrorReceived(Me, line)
                                                               End If
                                                           End Sub

            Try
                AddHandler process.OutputDataReceived, outputHandler
                AddHandler process.ErrorDataReceived, errorHandler

                process.Start()
                process.BeginOutputReadLine()
                process.BeginErrorReadLine()
                trackKey = TrackProcess(process)

                ' Enforce a hard timeout so hung executables cannot hold the task lock forever
                Using cts As New CancellationTokenSource(timeout)
                    Try
                        Await process.WaitForExitAsync(cts.Token)
                    Catch ex As OperationCanceledException
                        _logger.LogError($"Process exceeded {timeout.TotalMinutes:F0}-minute timeout. Killing process tree (PID {process.Id}).")
                        Try
                            process.Kill(entireProcessTree:=True)
                        Catch killEx As Exception
                            _logger.LogWarning($"Error killing timed-out process: {killEx.Message}")
                        End Try
                        Return False
                    End Try
                End Using

                If process.ExitCode <> 0 Then
                    _logger.LogError($"Process failed with exit code: {process.ExitCode}")
                    _logger.LogError($"Error output: {errors.ToString()}")
                    Return False
                End If

                Return True

            Catch ex As Exception
                _logger.LogError($"Error running process: {ex.Message}")
                Return False
            Finally
                RemoveHandler process.OutputDataReceived, outputHandler
                RemoveHandler process.ErrorDataReceived, errorHandler

                ' Kill the process if it's still running to prevent orphaned processes
                Try
                    If Not process.HasExited Then
                        process.Kill(entireProcessTree:=True)
                    End If
                Catch
                    ' Suppress errors during cleanup
                End Try

                UntrackProcess(trackKey, process)
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
        Implements IDisposable
        Public Property Success As Boolean
        Public Property Message As String
        Public Property Data As DataTable
        Public Property RowsAffected As Integer

        Public Sub Dispose() Implements IDisposable.Dispose
            Data?.Dispose()
            Data = Nothing
        End Sub
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
