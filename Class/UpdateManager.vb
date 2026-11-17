' -----------------------------------------------------------------------------
' LiteTask - part of LiteSuite. Released under The Unlicense (see LICENSE).
' File:    UpdateManager.vb
' Author:  LiteTask contributors
' Date:    2026-11-17
' Purpose: Self-update of LiteTask from its GitHub releases. The check, the
'          download (github.com releases of svtica/LiteTask only), the safe
'          extraction, the signature of LiteTask.exe and the replacement script
'          (services that run LiteTask.exe stopped and restarted, files of the
'          user kept, rollback if a copy fails, only the copy elevated) are done
'          by LiteSuite.Common's SelfUpdater. LiteTask keeps its own steps between
'          the preparation and the replacement: signature of the post-build
'          script, .NET runtime check and the post-build script itself.
' -----------------------------------------------------------------------------
Imports System.Drawing
Imports System.IO
Imports System.ServiceProcess

Namespace LiteTask
    Public Class UpdateManager
        Private ReadOnly _logger As Logger
        Private ReadOnly _tempPath As String

        Public Const RepositoryName As String = "LiteTask"
        Public Const ExecutableName As String = "LiteTask.exe"
        ''' <summary>Release asset built by Release.yml (Compress-Archive ... bin\LiteTask.zip).</summary>
        Public Const PackageName As String = "LiteTask.zip"
        Public Const PackagePattern As String = "^LiteTask.*\.zip$"
        Private Const PostBuildScriptName As String = "LiteTask-Post-Build.ps1"
        Private Const ServiceName As String = "LiteTaskService"

        ''' <summary>
        ''' SHA-1 thumbprint of the publisher certificate (res\certificate.cer) that the Release
        ''' workflow signs LiteTask.exe and LiteTask-Post-Build.ps1 with: the LiteSuite publisher.
        ''' SelfUpdater checks LiteTask.exe against it (or the signer of the running executable);
        ''' the post-build script is checked against the same thumbprints here.
        ''' </summary>
        Public Const ExpectedPublisherThumbprint As String = SelfUpdaterOptions.LiteSuiteThumbprint

        ''' <summary>
        ''' Files of the user kept by the replacement script when they already exist (PowerShell
        ''' -like patterns relative to the installation folder, '*' also matches '\'). LiteTask
        ''' writes its data under LiteTaskData\ (settings.xml, backup, logs, tools, modules,
        ''' notifications, temp) and, as fallbacks, *.log and the legacy *Credentials.dat next to
        ''' LiteTask.exe. None of them is in the release package today (it only ships
        ''' LiteTaskData\InstallModules.ps1 and LiteTaskData\lang\fr.xml, which must be updated),
        ''' and the script never deletes anything; these patterns only protect them should a
        ''' future package ship one of them.
        ''' </summary>
        Public Shared ReadOnly Property UserFilePatterns As IReadOnlyList(Of String) = New List(Of String) From {
            "LiteTaskData\settings.xml",
            "LiteTaskData\backup\*",
            "LiteTaskData\logs\*",
            "LiteTaskData\tools\*",
            "LiteTaskData\modules\*",
            "LiteTaskData\notifications\*",
            "ServiceCredentials.dat",
            "StoredCredentials.dat",
            "*.log"
        }

        Public Sub New(logger As Logger)
            _logger = logger
            _tempPath = Path.Combine(Application.StartupPath, "LiteTaskData", "temp")
        End Sub

        ''' <summary>The options of SelfUpdater for LiteTask.</summary>
        Public Shared Function CreateOptions(currentVersion As String, currentExecutablePath As String, tempDirectory As String) As SelfUpdaterOptions
            Dim options As New SelfUpdaterOptions With {
                .ApplicationName = "LiteTask",
                .Repository = RepositoryName,
                .PackagePattern = PackagePattern,
                .PackageName = PackageName,
                .ExecutableName = ExecutableName,
                .CurrentVersion = currentVersion,
                .CurrentExecutablePath = currentExecutablePath,
                .PinnedThumbprint = ExpectedPublisherThumbprint,
                .UserAgent = "LiteTask-Updater",
                .TempDirectory = tempDirectory
            }
            options.PreservePatterns.AddRange(UserFilePatterns)
            Return options
        End Function

        Private Function CreateUpdater() As SelfUpdater
            Return New SelfUpdater(CreateOptions(Application.ProductVersion, GetRunningExecutablePath(), _tempPath))
        End Function

        Private Shared Function GetRunningExecutablePath() As String
            Dim exePath = Environment.ProcessPath
            If String.IsNullOrWhiteSpace(exePath) Then exePath = Application.ExecutablePath
            Return exePath
        End Function

        ''' <summary>
        ''' Reads the latest release of svtica/LiteTask and returns version info.
        ''' </summary>
        Public Async Function CheckForUpdateAsync() As Task(Of UpdateInfo)
            Try
                _logger?.LogInfo("Checking for updates...")
                Dim check As UpdateCheckResult
                Using updater = CreateUpdater()
                    check = Await updater.CheckAsync()
                End Using
                Dim info = ToUpdateInfo(check)
                _logger?.LogInfo($"Update check complete. Current: {info.CurrentVersion}, Remote: {info.RemoteVersion}, Update available: {info.IsUpdateAvailable}, Package: {If(info.AssetName, "<none>")}")
                Return info
            Catch ex As Exception
                _logger?.LogError($"Error checking for updates: {ex.Message}")
                Throw
            End Try
        End Function

        ''' <summary>
        ''' UpdateInfo shown by MainForm from the result of SelfUpdater.CheckAsync. IsUpdateAvailable
        ''' is true for a newer release even without a package (MainForm then says so).
        ''' </summary>
        Public Shared Function ToUpdateInfo(check As UpdateCheckResult) As UpdateInfo
            If check Is Nothing Then Throw New ArgumentNullException(NameOf(check))
            Dim release = check.Release
            Dim info As New UpdateInfo With {
                .CurrentVersion = check.CurrentVersion,
                .RemoteVersion = check.LatestVersion,
                .IsUpdateAvailable = ReleaseVersion.IsNewer(check.CurrentVersion, check.LatestVersion),
                .ReleaseNotes = If(release?.Body, ""),
                .ReleaseName = release?.DisplayName,
                .Check = check
            }
            If check.Package IsNot Nothing Then
                info.DownloadUrl = check.Package.BrowserDownloadUrl
                info.AssetName = check.Package.Name
            End If
            Return info
        End Function

        ''' <summary>
        ''' Downloads, extracts and verifies the package (SelfUpdater), checks the post-build script
        ''' and the .NET runtime, runs the post-build script, then starts the replacement script and
        ''' exits. Reports live progress to an UpdateProgressForm. Nothing of the installation is
        ''' replaced before the replacement script runs.
        ''' </summary>
        Public Async Function DownloadAndApplyUpdateAsync(updateInfo As UpdateInfo, progressForm As UpdateProgressForm) As Task
            Dim update As PreparedUpdate = Nothing
            Dim applied = False

            Try
                If updateInfo Is Nothing OrElse updateInfo.Check Is Nothing OrElse updateInfo.Check.Package Is Nothing Then
                    Throw New UpdateException(UpdateFailure.NoPackage, "The release has no LiteTask package.")
                End If

                Using updater = CreateUpdater()
                    ' Steps 1 to 3: download, extraction, signature of LiteTask.exe (SelfUpdater)
                    Dim stageProgress As New Progress(Of UpdateStage)(Sub(stage) ReportStage(stage, progressForm))
                    _logger?.LogInfo($"Downloading update from {updateInfo.Check.Package.BrowserDownloadUrl}")
                    update = Await updater.PrepareAsync(updateInfo.Check, stageProgress)
                    _logger?.LogInfo($"Update {update.Version} extracted to {update.SourceDirectory}; signature verdict: {update.Verdict} ({update.Signature}).")

                    Dim expected = GetExpectedThumbprints()
                    If update.IsUnsigned Then
                        _logger?.LogWarning("Update package is unsigned; accepted because the running executable is an unsigned development build.")
                        progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                            "Update.Progress.UnsignedAllowed", "WARNING: the update package is not signed. Proceeding only because the running LiteTask.exe is itself an unsigned development build."), Color.Yellow)
                    Else
                        _logger?.LogInfo($"Update package signature accepted (signer {update.Signature?.SignerSubject}, thumbprint {update.Signature?.SignerThumbprint}).")
                        progressForm.AppendLog(String.Format(TranslationManager.Instance.GetTranslation(
                            "Update.Progress.SignatureValid", "Signature valid: signed by {0} (thumbprint {1})."),
                            update.Signature?.SignerSubject, update.Signature?.SignerThumbprint), Color.LimeGreen)
                    End If

                    ' Step 3b: the post-build script runs with -ExecutionPolicy Bypass before any file
                    ' is replaced, so in a signed package it must come from the same publisher as
                    ' LiteTask.exe. Nothing in an unsigned package can be pinned.
                    Dim postBuildScript = Path.Combine(update.SourceDirectory, PostBuildScriptName)
                    If Not File.Exists(postBuildScript) Then postBuildScript = Nothing
                    If postBuildScript IsNot Nothing AndAlso Not update.IsUnsigned AndAlso
                       Not IsScriptSignedByPublisher(postBuildScript, expected) Then
                        progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                            "Update.Progress.ScriptUnsigned", "The post-build script in the update package is not signed by the expected publisher."), Color.OrangeRed)
                        Abort(progressForm, update)
                        Return
                    End If
                    progressForm.SetProgress(55)

                    ' Step 3c: the package must be able to run on this machine. A release built on a
                    ' newer .NET major version needs its desktop runtime installed first; replacing
                    ' the files without it would leave LiteTask and its service unable to start.
                    Dim missingFrameworks = RuntimeRequirement.FindMissingFrameworks(update.SourceDirectory, ExecutableName)
                    If missingFrameworks.Count > 0 Then
                        Dim missingText = String.Join(", ", missingFrameworks.Select(Function(f) f.ToString()))
                        _logger?.LogError($"Update package requires {missingText}, which is not installed. No file was replaced.")
                        progressForm.AppendLog(String.Format(TranslationManager.Instance.GetTranslation(
                            "Update.Progress.RuntimeMissing",
                            "Update aborted: this version requires {0}, which is not installed on this computer. Install the .NET {1} Desktop Runtime (x64) from https://dotnet.microsoft.com/download, then run the update again. No file was replaced."),
                            missingText, missingFrameworks(0).Version.Major), Color.OrangeRed)
                        progressForm.SetCompleted(False)
                        update.Discard()
                        Return
                    End If

                    ' Step 4: post-build script on the extracted package, with live output
                    If postBuildScript IsNot Nothing Then
                        progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                            "Update.Progress.PostBuild", "Running post-build cleanup..."), Color.White)
                        _logger?.LogInfo($"Running PostBuild script: {postBuildScript}")
                        Await RunPostBuildScriptAsync(postBuildScript, progressForm)
                        _logger?.LogInfo("PostBuild script completed.")
                        progressForm.SetProgress(70)
                        progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                            "Update.Progress.PostBuildDone", "Post-build cleanup complete."), Color.LimeGreen)
                    Else
                        _logger?.LogInfo("No PostBuild script found, skipping.")
                        progressForm.SetProgress(70)
                        progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                            "Update.Progress.PostBuildSkipped", "No post-build script found, skipping."), Color.Gray)
                    End If

                    ' Step 5: the replacement script stops and restarts the services that run this
                    ' LiteTask.exe (found from their ImagePath). LiteTaskService installed with
                    ' another command line (not found that way: e.g. LiteTask.dll registered by the
                    ' old -register/InstallUtil path) is stopped here, as before.
                    Dim services = updater.FindServices()
                    _logger?.LogInfo($"Services that run {ExecutableName}: {If(services.Count = 0, "<none>", String.Join(", ", services))}")
                    If Not services.Any(Function(name) String.Equals(name, ServiceName, StringComparison.OrdinalIgnoreCase)) Then
                        StopUnmatchedServiceIfRunning(progressForm)
                    End If
                    progressForm.SetProgress(85)

                    ' Step 6: start the replacement script and exit
                    progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                        "Update.Progress.Launching", "Launching file replacement script..."), Color.White)
                    updater.Apply(update, Environment.ProcessId)
                    applied = True
                    _logger?.LogInfo($"Update script launched; log: {updater.LogPath}")
                    progressForm.SetProgress(95)
                End Using

                ' Mark completed before exiting
                progressForm.SetCompleted(True)

                ' Give the user a moment to read the final message, then exit
                Await Task.Delay(2000)
                _logger?.LogInfo("Exiting application for update...")
                Application.Exit()

            Catch ex As UpdateException
                _logger?.LogError($"Update failed ({ex.Failure}): {ex.Message}")
                progressForm.AppendLog(DescribeFailure(ex), Color.OrangeRed)
                If ex.Failure <> UpdateFailure.LaunchFailed Then
                    progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                        "Update.Progress.VerifyFailed", "Update aborted: no script was run and no file was replaced."), Color.OrangeRed)
                End If
                progressForm.SetCompleted(False)
                If Not applied Then update?.Discard()

            Catch ex As Exception
                _logger?.LogError($"Error applying update: {ex.Message}")
                progressForm.AppendLog($"Error: {ex.Message}", Color.OrangeRed)
                progressForm.SetCompleted(False)
                ' The work folder belongs to the replacement script once it is started
                If Not applied Then update?.Discard()
            End Try
        End Function

        ''' <summary>Shows a step reported by SelfUpdater with the existing Update.Progress.* texts.</summary>
        Private Sub ReportStage(stage As UpdateStage, progressForm As UpdateProgressForm)
            Select Case stage
                Case UpdateStage.Downloading
                    progressForm.SetProgress(5)
                    progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                        "Update.Progress.Downloading", "Downloading update..."), Color.White)
                Case UpdateStage.Extracting
                    _logger?.LogInfo("Download complete.")
                    progressForm.SetProgress(30)
                    progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                        "Update.Progress.DownloadDone", "Download complete."), Color.LimeGreen)
                    progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                        "Update.Progress.Extracting", "Extracting files..."), Color.White)
                Case UpdateStage.VerifyingSignature
                    _logger?.LogInfo("Extraction complete.")
                    progressForm.SetProgress(45)
                    progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                        "Update.Progress.ExtractDone", "Extraction complete."), Color.LimeGreen)
                    progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                        "Update.Progress.Verifying", "Verifying package signature..."), Color.White)
                Case UpdateStage.SignatureAccepted, UpdateStage.UnsignedAccepted
                    ' Shown once PrepareAsync returns, with the signer of the package
                    progressForm.SetProgress(50)
            End Select
        End Sub

        ''' <summary>Translated text of a failure of SelfUpdater.</summary>
        Private Shared Function DescribeFailure(ex As UpdateException) As String
            Dim tm = TranslationManager.Instance
            Select Case ex.Failure
                Case UpdateFailure.SignatureRejected
                    Return DescribeRejectedSignature(ex)
                Case UpdateFailure.ExecutableNotFound
                    Return tm.GetTranslation("Update.Progress.ExeNotFound", "LiteTask.exe was not found in the update package.")
                Case UpdateFailure.NoPackage
                    Return tm.GetTranslation("Update.NoPackage", "No download package found in the release. Please download manually from GitHub.")
                Case UpdateFailure.CheckFailed
                    Return tm.GetTranslation("Update.CheckFailed", "Update check failed.") & " " & ex.Message
                Case UpdateFailure.UnsafeUrl
                    Return tm.GetTranslation("Update.Progress.UnsafeUrl", "The download link of the package is not a release of svtica/LiteTask on github.com.")
                Case UpdateFailure.DownloadFailed
                    Return String.Format(tm.GetTranslation("Update.Progress.DownloadFailed", "The package could not be downloaded: {0}"), ex.Message)
                Case UpdateFailure.InvalidPackage
                    Return String.Format(tm.GetTranslation("Update.Progress.InvalidPackage", "The package is invalid and was not extracted: {0}"), ex.Message)
                Case UpdateFailure.LaunchFailed
                    Return String.Format(tm.GetTranslation("Update.Progress.LaunchFailed", "The file replacement script could not be started: {0}"), ex.Message)
                Case Else
                    Return $"Error: {ex.Message}"
            End Select
        End Function

        Private Shared Function DescribeRejectedSignature(ex As UpdateException) As String
            Dim tm = TranslationManager.Instance
            Dim signature = ex.Signature
            Dim subject = signature?.SignerSubject
            Dim thumbprint = signature?.SignerThumbprint
            Dim errorCode = If(signature Is Nothing, 0, signature.ErrorCode)
            Dim verdict = If(ex.Verdict.HasValue, ex.Verdict.Value, UpdateSignatureVerdict.RejectedInvalidSignature)
            Select Case verdict
                Case UpdateSignatureVerdict.RejectedUnsigned
                    Return tm.GetTranslation("Update.Progress.UnsignedRejected", "The update package is not signed by the expected publisher.")
                Case UpdateSignatureVerdict.RejectedTampered
                    Return String.Format(tm.GetTranslation("Update.Progress.SignatureTampered",
                        "The signature of {0} does not match its content (file tampered or corrupted)."), ExecutableName)
                Case UpdateSignatureVerdict.RejectedPublisherMismatch
                    Return String.Format(tm.GetTranslation("Update.Progress.PublisherMismatch",
                        "{0} is signed by an unexpected publisher ({1}, thumbprint {2})."), ExecutableName, subject, thumbprint)
                Case Else
                    Return String.Format(tm.GetTranslation("Update.Progress.SignatureInvalid",
                        "The signature of {0} could not be validated (error 0x{1})."), ExecutableName, errorCode.ToString("X8"))
            End Select
        End Function

        ''' <summary>Logs the abort of a verified package, discards it and marks the form failed.</summary>
        Private Sub Abort(progressForm As UpdateProgressForm, update As PreparedUpdate)
            _logger?.LogError("Update package rejected. No script was run and no file was replaced.")
            progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                "Update.Progress.VerifyFailed", "Update aborted: no script was run and no file was replaced."), Color.OrangeRed)
            progressForm.SetCompleted(False)
            update?.Discard()
        End Sub

        ''' <summary>
        ''' The thumbprints a signed package must carry: the pinned publisher plus the signer of the
        ''' running LiteTask.exe (same rule as SelfUpdater for LiteTask.exe).
        ''' </summary>
        Private Function GetExpectedThumbprints() As List(Of String)
            Return UpdateSignaturePolicy.ExpectedThumbprints(ExpectedPublisherThumbprint, GetRunningExecutableSignature())
        End Function

        Private Function IsScriptSignedByPublisher(scriptPath As String, expected As IEnumerable(Of String)) As Boolean
            Dim script = AuthenticodeVerifier.Verify(scriptPath)
            _logger?.LogInfo($"Signature of {scriptPath}: {script}")
            If Not script.IsSignatureValid OrElse Not AuthenticodeVerifier.ThumbprintMatches(script.SignerThumbprint, expected) Then
                _logger?.LogError($"Update package rejected: {PostBuildScriptName} is not validly signed by the expected publisher ({script}); expected one of [{String.Join(", ", expected)}].")
                Return False
            End If
            _logger?.LogInfo($"{PostBuildScriptName} signature accepted.")
            Return True
        End Function

        ''' <summary>
        ''' Authenticode result of the currently running LiteTask.exe, or Nothing when it cannot be determined.
        ''' </summary>
        Private Function GetRunningExecutableSignature() As AuthenticodeResult
            Try
                Dim exePath = GetRunningExecutablePath()
                If String.IsNullOrWhiteSpace(exePath) OrElse Not File.Exists(exePath) Then Return Nothing
                Return AuthenticodeVerifier.Verify(exePath)
            Catch ex As Exception
                _logger?.LogWarning($"Could not determine the signature of the running executable: {ex.Message}")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Runs the PostBuild script asynchronously, streaming output line-by-line to the progress form.
        ''' </summary>
        Private Async Function RunPostBuildScriptAsync(scriptPath As String, progressForm As UpdateProgressForm) As Task
            Try
                Dim scriptDir = Path.GetDirectoryName(scriptPath)
                Dim startInfo As New ProcessStartInfo() With {
                    .FileName = "powershell.exe",
                    .Arguments = $"-NoProfile -ExecutionPolicy Bypass -File ""{scriptPath}""",
                    .WorkingDirectory = scriptDir,
                    .UseShellExecute = False,
                    .CreateNoWindow = True,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True
                }

                Using proc As New Process With {.StartInfo = startInfo}
                    AddHandler proc.OutputDataReceived, Sub(s, args)
                                                            If args.Data IsNot Nothing Then
                                                                _logger?.LogInfo($"PostBuild: {args.Data}")
                                                                progressForm.AppendLog("  " & args.Data, Color.LightGray)
                                                            End If
                                                        End Sub

                    AddHandler proc.ErrorDataReceived, Sub(s, args)
                                                           If args.Data IsNot Nothing Then
                                                               _logger?.LogError($"PostBuild error: {args.Data}")
                                                               progressForm.AppendLog("  " & args.Data, Color.OrangeRed)
                                                           End If
                                                       End Sub

                    proc.Start()
                    proc.BeginOutputReadLine()
                    proc.BeginErrorReadLine()

                    ' Wait asynchronously with a 60-second timeout
                    Using cts As New Threading.CancellationTokenSource(60000)
                        Try
                            Await proc.WaitForExitAsync(cts.Token)
                        Catch ex As OperationCanceledException
                            proc.Kill()
                            _logger?.LogError("PostBuild script timed out after 60 seconds.")
                            progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                                "Update.Progress.PostBuildTimeout", "Post-build script timed out."), Color.OrangeRed)
                        End Try
                    End Using
                End Using
            Catch ex As Exception
                _logger?.LogError($"Error running PostBuild script: {ex.Message}")
                progressForm.AppendLog($"PostBuild error: {ex.Message}", Color.OrangeRed)
            End Try
        End Function

        ''' <summary>
        ''' LiteTaskService whose ImagePath does not resolve to this LiteTask.exe (installed with
        ''' another path form or by another means) is not stopped by the replacement script: stop it
        ''' here as before, so that the copy does not fail on locked files. The script does not
        ''' restart it; the user is told to reinstall it from the Service menu (or -unregister then
        ''' -register), which registers "LiteTask.exe" -service so that later updates find it.
        ''' </summary>
        Private Sub StopUnmatchedServiceIfRunning(progressForm As UpdateProgressForm)
            If Not ServiceLocator.IsActive(ServiceName) Then Return
            Try
                Using sc = New ServiceController(ServiceName)
                    _logger?.LogWarning($"Service {ServiceName} is running but its ImagePath does not point to this {ExecutableName}; stopping it before the update.")
                    progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                        "Update.Progress.StoppingService", "Stopping service..."), Color.White)
                    If sc.Status <> ServiceControllerStatus.Stopped AndAlso sc.Status <> ServiceControllerStatus.StopPending Then
                        sc.Stop()
                    End If
                    sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(30))
                    _logger?.LogInfo($"Service {ServiceName} stopped.")
                    progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                        "Update.Progress.ServiceStopped", "Service stopped."), Color.LimeGreen)
                    progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                        "Update.Progress.ServiceManualRestart", "LiteTaskService does not run this LiteTask.exe: after the update, reinstall it (menu Service > Uninstall Service, then Install Service) and start it."), Color.Yellow)
                End Using
            Catch ex As Exception
                _logger?.LogError($"Error stopping service: {ex.Message}")
                progressForm.AppendLog($"Service error: {ex.Message}", Color.OrangeRed)
            End Try
        End Sub

        Public Sub Dispose()
            ' Nothing held: each operation disposes its SelfUpdater
        End Sub
    End Class

    Public Class UpdateInfo
        Public Property CurrentVersion As String
        Public Property RemoteVersion As String
        Public Property IsUpdateAvailable As Boolean
        Public Property DownloadUrl As String
        Public Property AssetName As String
        Public Property ReleaseNotes As String
        Public Property ReleaseName As String
        ''' <summary>Result of SelfUpdater.CheckAsync, passed back to DownloadAndApplyUpdateAsync.</summary>
        Public Property Check As UpdateCheckResult
    End Class
End Namespace
