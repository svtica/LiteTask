' -----------------------------------------------------------------------------
' LiteTask - part of LiteSuite. Released under The Unlicense (see LICENSE).
' File:    ServiceRegistration.vb
' Author:  LiteTask contributors
' Date:    2026-11-17
' Purpose: Single implementation of the LiteTaskService registration, shared by
'          the "Install service" / "Uninstall service" menu items and the
'          -register / -unregister command-line options. The service runs
'          "<LiteTask.exe>" -service (never LiteTask.dll: under .NET 10 the
'          executing assembly is the DLL, not a service executable), so the
'          self-updater, which finds services by their ImagePath, recognises it.
'          The pure Build* functions are pinned by ServiceRegistrationTests.
' -----------------------------------------------------------------------------
Namespace LiteTask
    Public NotInheritable Class ServiceRegistration

        Public Const ServiceName As String = "LiteTaskService"
        Public Const ServiceArgument As String = "-service"

        Private Sub New()
        End Sub

        ''' <summary>
        ''' Full path of the running LiteTask.exe: Environment.ProcessPath, falling back
        ''' to Application.ExecutablePath. Never Assembly.Location (LiteTask.dll).
        ''' </summary>
        Public Shared Function GetExecutablePath() As String
            Dim processPath As String = Environment.ProcessPath
            If String.IsNullOrEmpty(processPath) Then
                processPath = Application.ExecutablePath
            End If
            Return processPath
        End Function

        ''' <summary>
        ''' Service command line as the SCM stores it (ImagePath):
        ''' "C:\Program Files\LiteTask\LiteTask.exe" -service. The path is always quoted
        ''' so that a path with spaces is not ambiguous.
        ''' </summary>
        Public Shared Function BuildImagePath(exePath As String) As String
            Return $"""{exePath}"" {ServiceArgument}"
        End Function

        ''' <summary>
        ''' sc.exe arguments that create the service, same options as the menu has always
        ''' used (automatic start, default LocalSystem account):
        ''' create LiteTaskService binPath= "\"C:\...\LiteTask.exe\" -service" start= auto
        ''' The inner quotes are escaped with a backslash so that sc.exe receives the
        ''' quoted ImagePath as a single binPath value. A file path never ends with a
        ''' backslash, so \" is always an escaped quote here.
        ''' </summary>
        Public Shared Function BuildCreateArguments(exePath As String) As String
            Dim escapedImagePath As String = BuildImagePath(exePath).Replace("""", "\""")
            Return $"create {ServiceName} binPath= ""{escapedImagePath}"" start= auto"
        End Function

        ''' <summary>sc.exe arguments that delete the service.</summary>
        Public Shared Function BuildDeleteArguments() As String
            Return $"delete {ServiceName}"
        End Function

        ''' <summary>Registers LiteTaskService for this LiteTask.exe. Throws on failure.</summary>
        Public Shared Sub Install(logPath As String, Optional logger As Logger = Nothing)
            RunSc(BuildCreateArguments(GetExecutablePath()), "Service Installation", logPath, logger)
        End Sub

        ''' <summary>Deletes LiteTaskService. Throws on failure.</summary>
        Public Shared Sub Uninstall(logPath As String, Optional logger As Logger = Nothing)
            RunSc(BuildDeleteArguments(), "Service Uninstallation", logPath, logger)
        End Sub

        ''' <summary>
        ''' Runs %SystemRoot%\System32\sc.exe directly (no cmd.exe, so no shell
        ''' re-parsing of the quotes), logs the command and its output to
        ''' <paramref name="logPath"/> when given, and throws when sc.exe fails.
        ''' sc.exe writes its errors to standard output, so they are included.
        ''' </summary>
        Public Shared Sub RunSc(arguments As String, operationType As String, logPath As String, Optional logger As Logger = Nothing)
            Dim scPath As String = Path.Combine(Environment.SystemDirectory, "sc.exe")
            AppendLog(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {operationType}{Environment.NewLine}" &
                               $"Command: sc {arguments}{Environment.NewLine}")

            Dim startInfo As New ProcessStartInfo With {
                .FileName = scPath,
                .Arguments = arguments,
                .UseShellExecute = False,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True,
                .CreateNoWindow = True
            }

            Using proc As New Process With {.StartInfo = startInfo}
                proc.Start()
                Dim outputTask = proc.StandardOutput.ReadToEndAsync()
                Dim errors As String = proc.StandardError.ReadToEnd()
                Dim output As String = outputTask.Result
                proc.WaitForExit()

                Dim entry As String = $"Output: {output}{Environment.NewLine}"
                If Not String.IsNullOrEmpty(errors) Then
                    entry &= $"Errors: {errors}{Environment.NewLine}"
                End If
                entry &= $"Exit Code: {proc.ExitCode}{Environment.NewLine}" & New String("-"c, 50) & Environment.NewLine
                AppendLog(logPath, entry)

                If proc.ExitCode <> 0 Then
                    Dim detail As String = If(String.IsNullOrWhiteSpace(errors), output, errors).Trim()
                    logger?.LogError($"sc {arguments} failed with exit code {proc.ExitCode}: {detail}")
                    Throw New Exception($"sc.exe failed with exit code {proc.ExitCode}. {detail}")
                End If

                logger?.LogInfo($"Service command executed successfully: sc {arguments}")
            End Using
        End Sub

        Private Shared Sub AppendLog(logPath As String, text As String)
            If String.IsNullOrEmpty(logPath) Then Return
            Try
                Dim directoryPath As String = Path.GetDirectoryName(logPath)
                If Not String.IsNullOrEmpty(directoryPath) Then
                    Directory.CreateDirectory(directoryPath)
                End If
                File.AppendAllText(logPath, text)
            Catch
                ' Logging must never block the service operation.
            End Try
        End Sub

    End Class
End Namespace
