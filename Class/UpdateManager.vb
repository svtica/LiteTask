Imports System.IO.Compression
Imports System.Net.Http
Imports System.Text.Json
Imports System.ServiceProcess

Namespace LiteTask
    Public Class UpdateManager
        Private ReadOnly _logger As Logger
        Private ReadOnly _httpClient As HttpClient
        Private ReadOnly _tempPath As String
        Private ReadOnly _appPath As String

        Private Const GitHubApiUrl As String = "https://api.github.com/repos/svtica/LiteTask/releases/latest"
        Private Const ServiceName As String = "LiteTaskService"
        Private Const ExecutableName As String = "LiteTask.exe"
        Private Const PostBuildScriptName As String = "LiteTask-Post-Build.ps1"

        ''' <summary>
        ''' SHA-1 thumbprint of the publisher certificate (res\certificate.cer, CN=SVTI.ca) that the
        ''' Release workflow signs LiteTask.exe with. An update is only applied when the extracted
        ''' LiteTask.exe carries a valid Authenticode signature from this certificate (or from the
        ''' certificate that signed the running executable). Rotating the signing certificate
        ''' requires updating this constant in a release signed with the previous certificate.
        ''' </summary>
        Public Const ExpectedPublisherThumbprint As String = "EC2B412C4025C2F1C8DB512CEC66632E65C017F0"

        Public Sub New(logger As Logger)
            _logger = logger
            _appPath = Application.StartupPath
            _tempPath = Path.Combine(_appPath, "LiteTaskData", "temp")
            _httpClient = New HttpClient()
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "LiteTask-Updater")
        End Sub

        ''' <summary>
        ''' Checks GitHub for the latest release version and returns version info.
        ''' </summary>
        Public Async Function CheckForUpdateAsync() As Task(Of UpdateInfo)
            Try
                _logger?.LogInfo("Checking for updates...")
                Dim response = Await _httpClient.GetAsync(GitHubApiUrl)
                response.EnsureSuccessStatusCode()

                Dim json = Await response.Content.ReadAsStringAsync()
                Dim info As UpdateInfo
                Using doc = JsonDocument.Parse(json)
                    Dim root = doc.RootElement

                    Dim tagName = root.GetProperty("tag_name").GetString()
                    Dim remoteVersion = tagName.TrimStart("v"c)
                    Dim currentVersion = Application.ProductVersion

                    info = New UpdateInfo With {
                        .CurrentVersion = currentVersion,
                        .RemoteVersion = remoteVersion,
                        .IsUpdateAvailable = IsNewerVersion(currentVersion, remoteVersion),
                        .ReleaseNotes = If(root.TryGetProperty("body", Nothing), root.GetProperty("body").GetString(), ""),
                        .ReleaseName = If(root.TryGetProperty("name", Nothing), root.GetProperty("name").GetString(), tagName)
                    }

                    ' Find the zip asset download URL. Only accept HTTPS links so a
                    ' tampered API answer cannot redirect the download to plain HTTP.
                    If root.TryGetProperty("assets", Nothing) Then
                        For Each asset In root.GetProperty("assets").EnumerateArray()
                            Dim assetName = asset.GetProperty("name").GetString()
                            Dim downloadUrl = asset.GetProperty("browser_download_url").GetString()
                            If assetName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase) AndAlso
                               downloadUrl IsNot Nothing AndAlso
                               downloadUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
                                info.DownloadUrl = downloadUrl
                                ' Never let the remote asset name escape the temp folder.
                                info.AssetName = Path.GetFileName(assetName)
                                Exit For
                            End If
                        Next
                    End If
                End Using

                _logger?.LogInfo($"Update check complete. Current: {info.CurrentVersion}, Remote: {info.RemoteVersion}, Update available: {info.IsUpdateAvailable}")
                Return info

            Catch ex As Exception
                _logger?.LogError($"Error checking for updates: {ex.Message}")
                Throw
            End Try
        End Function

        ''' <summary>
        ''' Downloads, extracts, runs PostBuild, stops service, and launches the update script.
        ''' Reports live progress to an UpdateProgressForm.
        ''' </summary>
        Public Async Function DownloadAndApplyUpdateAsync(updateInfo As UpdateInfo, progressForm As UpdateProgressForm) As Task
            Dim zipPath As String = Nothing
            Dim extractPath As String = Nothing

            Try
                ' Only accept HTTPS download links from the release API.
                Dim downloadUri As Uri = Nothing
                If Not Uri.TryCreate(updateInfo.DownloadUrl, UriKind.Absolute, downloadUri) OrElse
                   Not String.Equals(downloadUri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) Then
                    Throw New InvalidOperationException($"Refusing to download update from non-HTTPS URL: {updateInfo.DownloadUrl}")
                End If

                Directory.CreateDirectory(_tempPath)
                Dim assetFileName = Path.GetFileName(If(updateInfo.AssetName, String.Empty))
                zipPath = Path.Combine(_tempPath, If(String.IsNullOrWhiteSpace(assetFileName), "LiteTask-update.zip", assetFileName))
                extractPath = Path.Combine(_tempPath, "update")

                ' Clean previous update attempt
                If Directory.Exists(extractPath) Then
                    Directory.Delete(extractPath, True)
                End If

                ' Step 1: Download
                progressForm.SetProgress(5)
                progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                    "Update.Progress.Downloading", "Downloading update..."), Drawing.Color.White)
                _logger?.LogInfo($"Downloading update from {updateInfo.DownloadUrl}")
                Await DownloadFileAsync(updateInfo.DownloadUrl, zipPath)
                _logger?.LogInfo("Download complete.")
                progressForm.SetProgress(30)
                progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                    "Update.Progress.DownloadDone", "Download complete."), Drawing.Color.LimeGreen)

                ' Step 2: Extract
                progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                    "Update.Progress.Extracting", "Extracting files..."), Drawing.Color.White)
                _logger?.LogInfo($"Extracting to {extractPath}")
                If Directory.Exists(extractPath) Then
                    Directory.Delete(extractPath, True)
                End If
                ZipFile.ExtractToDirectory(zipPath, extractPath)
                _logger?.LogInfo("Extraction complete.")
                progressForm.SetProgress(50)
                progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                    "Update.Progress.ExtractDone", "Extraction complete."), Drawing.Color.LimeGreen)

                ' Step 3: Verify the Authenticode signature of the extracted package.
                ' Nothing from the archive is executed and no file is replaced until this passes.
                progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                    "Update.Progress.Verifying", "Verifying package signature..."), Drawing.Color.White)
                Dim postBuildScript = FindFile(extractPath, PostBuildScriptName)
                If Not VerifyUpdatePackage(extractPath, postBuildScript, progressForm) Then
                    _logger?.LogError("Update package rejected by signature verification. No script was run and no file was replaced.")
                    progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                        "Update.Progress.VerifyFailed", "Update aborted: no script was run and no file was replaced."), Drawing.Color.OrangeRed)
                    progressForm.SetCompleted(False)
                    CleanupTempFiles(zipPath, extractPath)
                    Return
                End If
                progressForm.SetProgress(55)

                ' Step 4: Run PostBuild script with live output
                If postBuildScript IsNot Nothing Then
                    progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                        "Update.Progress.PostBuild", "Running post-build cleanup..."), Drawing.Color.White)
                    _logger?.LogInfo($"Running PostBuild script: {postBuildScript}")
                    Await RunPostBuildScriptAsync(postBuildScript, progressForm)
                    _logger?.LogInfo("PostBuild script completed.")
                    progressForm.SetProgress(70)
                    progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                        "Update.Progress.PostBuildDone", "Post-build cleanup complete."), Drawing.Color.LimeGreen)
                Else
                    _logger?.LogInfo("No PostBuild script found, skipping.")
                    progressForm.SetProgress(70)
                    progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                        "Update.Progress.PostBuildSkipped", "No post-build script found, skipping."), Drawing.Color.Gray)
                End If

                ' Step 5: Stop service if running
                progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                    "Update.Progress.StoppingService", "Stopping service..."), Drawing.Color.White)
                StopServiceIfRunning(progressForm)
                progressForm.SetProgress(85)

                ' Step 6: Launch update script and exit
                progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                    "Update.Progress.Launching", "Launching file replacement script..."), Drawing.Color.White)
                LaunchUpdateScript(extractPath)
                progressForm.SetProgress(95)

                ' Mark completed before exiting
                progressForm.SetCompleted(True)

                ' Give the user a moment to read the final message, then exit
                Await Task.Delay(2000)
                _logger?.LogInfo("Exiting application for update...")
                Application.Exit()

            Catch ex As Exception
                _logger?.LogError($"Error applying update: {ex.Message}")
                progressForm.AppendLog($"Error: {ex.Message}", Drawing.Color.OrangeRed)
                progressForm.SetCompleted(False)
                ' Clean up on error
                CleanupTempFiles(zipPath, extractPath)
            End Try
        End Function

        Private Async Function DownloadFileAsync(url As String, filePath As String) As Task
            Using response = Await _httpClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead)
                response.EnsureSuccessStatusCode()
                Using contentStream = Await response.Content.ReadAsStreamAsync()
                    Using fileStream As New FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, True)
                        Await contentStream.CopyToAsync(fileStream)
                    End Using
                End Using
            End Using
        End Function

        Private Function FindFile(rootPath As String, fileName As String) As String
            Try
                Dim files = Directory.GetFiles(rootPath, fileName, SearchOption.AllDirectories)
                Return If(files.Length > 0, files(0), Nothing)
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Locates LiteTask.exe in the extracted tree using the same rules as the update script:
        ''' archive root, then a single sub-folder, then the first match anywhere.
        ''' </summary>
        Private Function LocateUpdateExecutable(extractPath As String) As String
            Try
                Dim rootExe = Path.Combine(extractPath, ExecutableName)
                If File.Exists(rootExe) Then Return rootExe

                Dim subDirs = Directory.GetDirectories(extractPath)
                If subDirs.Length = 1 Then
                    Dim nestedExe = Path.Combine(subDirs(0), ExecutableName)
                    If File.Exists(nestedExe) Then Return nestedExe
                End If
            Catch
            End Try
            Return FindFile(extractPath, ExecutableName)
        End Function

        ''' <summary>
        ''' Verifies that the extracted LiteTask.exe carries a valid Authenticode signature from the
        ''' expected publisher (see <see cref="ExpectedPublisherThumbprint"/>). Policy:
        '''  - signed by the expected publisher: accepted; the post-build script shipped in the
        '''    archive must then be signed by the same publisher, since it runs before any file is
        '''    replaced;
        '''  - unsigned package while the running LiteTask.exe is itself unsigned (development
        '''    build): accepted with a warning;
        '''  - unsigned, tampered, invalid or signed by another publisher: rejected.
        ''' Returns False when the update must not proceed. Reasons are logged and shown in the form.
        ''' </summary>
        Private Function VerifyUpdatePackage(extractPath As String, postBuildScript As String, progressForm As UpdateProgressForm) As Boolean
            Dim updateExe = LocateUpdateExecutable(extractPath)
            If updateExe Is Nothing Then
                _logger?.LogError($"{ExecutableName} not found in extracted update package at {extractPath}.")
                progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                    "Update.Progress.ExeNotFound", "LiteTask.exe was not found in the update package."), Drawing.Color.OrangeRed)
                Return False
            End If

            Dim running = GetRunningExecutableSignature()
            Dim expected = UpdateSignaturePolicy.ExpectedThumbprints(ExpectedPublisherThumbprint, running)
            Dim package = AuthenticodeVerifier.Verify(updateExe)
            _logger?.LogInfo($"Signature of running executable: {If(running Is Nothing, "<unknown>", running.ToString())}")
            _logger?.LogInfo($"Signature of {updateExe}: {package}")

            Dim verdict = UpdateSignaturePolicy.Evaluate(package, running, expected)
            Select Case verdict
                Case UpdateSignatureVerdict.Accepted
                    _logger?.LogInfo($"Update package signature accepted (signer {package.SignerSubject}, thumbprint {package.SignerThumbprint}).")
                    progressForm.AppendLog(String.Format(TranslationManager.Instance.GetTranslation(
                        "Update.Progress.SignatureValid", "Signature valid: signed by {0} (thumbprint {1})."),
                        package.SignerSubject, package.SignerThumbprint), Drawing.Color.LimeGreen)

                Case UpdateSignatureVerdict.AcceptedUnsignedDevBuild
                    _logger?.LogWarning("Update package is unsigned; accepted because the running executable is an unsigned development build.")
                    progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                        "Update.Progress.UnsignedAllowed", "WARNING: the update package is not signed. Proceeding only because the running LiteTask.exe is itself an unsigned development build."), Drawing.Color.Yellow)
                    ' Nothing in an unsigned package can be pinned, including the post-build script.
                    Return True

                Case UpdateSignatureVerdict.RejectedUnsigned
                    _logger?.LogError("Update package rejected: LiteTask.exe is not signed while the running executable is signed.")
                    progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                        "Update.Progress.UnsignedRejected", "The update package is not signed by the expected publisher."), Drawing.Color.OrangeRed)
                    Return False

                Case UpdateSignatureVerdict.RejectedTampered
                    _logger?.LogError($"Update package rejected: the signature of {updateExe} does not match its content (0x{package.ErrorCode:X8}).")
                    progressForm.AppendLog(String.Format(TranslationManager.Instance.GetTranslation(
                        "Update.Progress.SignatureTampered", "The signature of {0} does not match its content (file tampered or corrupted)."),
                        ExecutableName), Drawing.Color.OrangeRed)
                    Return False

                Case UpdateSignatureVerdict.RejectedPublisherMismatch
                    _logger?.LogError($"Update package rejected: {updateExe} is signed by an unexpected publisher ({package.SignerSubject}, thumbprint {package.SignerThumbprint}); expected one of [{String.Join(", ", expected)}].")
                    progressForm.AppendLog(String.Format(TranslationManager.Instance.GetTranslation(
                        "Update.Progress.PublisherMismatch", "{0} is signed by an unexpected publisher ({1}, thumbprint {2})."),
                        ExecutableName, package.SignerSubject, package.SignerThumbprint), Drawing.Color.OrangeRed)
                    Return False

                Case Else
                    _logger?.LogError($"Update package rejected: the signature of {updateExe} could not be validated (0x{package.ErrorCode:X8}).")
                    progressForm.AppendLog(String.Format(TranslationManager.Instance.GetTranslation(
                        "Update.Progress.SignatureInvalid", "The signature of {0} could not be validated (error 0x{1})."),
                        ExecutableName, package.ErrorCode.ToString("X8")), Drawing.Color.OrangeRed)
                    Return False
            End Select

            ' Signed release: the post-build script runs with -ExecutionPolicy Bypass before any file
            ' is replaced, so it must come from the same publisher as the executable.
            If postBuildScript IsNot Nothing Then
                Dim script = AuthenticodeVerifier.Verify(postBuildScript)
                _logger?.LogInfo($"Signature of {postBuildScript}: {script}")
                If Not script.IsSignatureValid OrElse Not AuthenticodeVerifier.ThumbprintMatches(script.SignerThumbprint, expected) Then
                    _logger?.LogError($"Update package rejected: {PostBuildScriptName} is not validly signed by the expected publisher ({script}).")
                    progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                        "Update.Progress.ScriptUnsigned", "The post-build script in the update package is not signed by the expected publisher."), Drawing.Color.OrangeRed)
                    Return False
                End If
                _logger?.LogInfo($"{PostBuildScriptName} signature accepted.")
            End If

            Return True
        End Function

        ''' <summary>
        ''' Authenticode result of the currently running LiteTask.exe, or Nothing when it cannot be determined.
        ''' </summary>
        Private Function GetRunningExecutableSignature() As AuthenticodeResult
            Try
                Dim exePath = Environment.ProcessPath
                If String.IsNullOrWhiteSpace(exePath) Then exePath = Application.ExecutablePath
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
                                                                progressForm.AppendLog("  " & args.Data, Drawing.Color.LightGray)
                                                            End If
                                                        End Sub

                    AddHandler proc.ErrorDataReceived, Sub(s, args)
                                                           If args.Data IsNot Nothing Then
                                                               _logger?.LogError($"PostBuild error: {args.Data}")
                                                               progressForm.AppendLog("  " & args.Data, Drawing.Color.OrangeRed)
                                                           End If
                                                       End Sub

                    proc.Start()
                    proc.BeginOutputReadLine()
                    proc.BeginErrorReadLine()

                    ' Wait asynchronously with a 60-second timeout
                    Dim cts As New Threading.CancellationTokenSource(60000)
                    Try
                        Await proc.WaitForExitAsync(cts.Token)
                    Catch ex As OperationCanceledException
                        proc.Kill()
                        _logger?.LogError("PostBuild script timed out after 60 seconds.")
                        progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                            "Update.Progress.PostBuildTimeout", "Post-build script timed out."), Drawing.Color.OrangeRed)
                    End Try
                End Using
            Catch ex As Exception
                _logger?.LogError($"Error running PostBuild script: {ex.Message}")
                progressForm.AppendLog($"PostBuild error: {ex.Message}", Drawing.Color.OrangeRed)
            End Try
        End Function

        Private Sub StopServiceIfRunning(progressForm As UpdateProgressForm)
            Try
                Using sc = New ServiceController(ServiceName)
                    If sc.Status = ServiceControllerStatus.Running OrElse
                       sc.Status = ServiceControllerStatus.StartPending Then
                        _logger?.LogInfo($"Stopping service {ServiceName}...")
                        sc.Stop()
                        sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(30))
                        _logger?.LogInfo($"Service {ServiceName} stopped.")
                        progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                            "Update.Progress.ServiceStopped", "Service stopped."), Drawing.Color.LimeGreen)
                    Else
                        _logger?.LogInfo($"Service {ServiceName} is not running (status: {sc.Status}).")
                        progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                            "Update.Progress.ServiceNotRunning", "Service is not running, skipping."), Drawing.Color.Gray)
                    End If
                End Using
            Catch ex As InvalidOperationException
                ' Service not installed
                _logger?.LogInfo($"Service {ServiceName} is not installed, skipping stop.")
                progressForm.AppendLog(TranslationManager.Instance.GetTranslation(
                    "Update.Progress.ServiceNotInstalled", "Service is not installed, skipping."), Drawing.Color.Gray)
            Catch ex As Exception
                _logger?.LogError($"Error stopping service: {ex.Message}")
                progressForm.AppendLog($"Service error: {ex.Message}", Drawing.Color.OrangeRed)
            End Try
        End Sub

        ''' <summary>
        ''' Creates and launches a PowerShell script that waits for the app to exit,
        ''' then replaces the files and restarts the application.
        ''' </summary>
        Private Sub LaunchUpdateScript(extractPath As String)
            Dim updateScriptPath = Path.Combine(_tempPath, "LiteTask-Update.ps1")
            Dim currentPid = Environment.ProcessId

            Dim scriptContent = $"
# LiteTask Update Script - Auto-generated
$processId = {currentPid}
$appPath = '{_appPath.Replace("'", "''")}'
$updateSource = '{extractPath.Replace("'", "''")}'

# Wait for the application to exit
try {{
    $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
    if ($process) {{
        $process.WaitForExit(30000) | Out-Null
        if (!$process.HasExited) {{
            $process | Stop-Process -Force
        }}
    }}
}} catch {{
    # Process already exited
}}

# Determine the source folder (handle nested folder structure)
$sourceFolder = $updateSource
$subDirs = Get-ChildItem -Path $updateSource -Directory
if ($subDirs.Count -eq 1 -and (Test-Path (Join-Path $subDirs[0].FullName 'LiteTask.exe'))) {{
    $sourceFolder = $subDirs[0].FullName
}} elseif (!(Test-Path (Join-Path $updateSource 'LiteTask.exe'))) {{
    $exeFile = Get-ChildItem -Path $updateSource -Filter 'LiteTask.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($exeFile) {{
        $sourceFolder = $exeFile.DirectoryName
    }}
}}

# Copy all files from update source to application directory
try {{
    Get-ChildItem -Path $sourceFolder -Recurse | ForEach-Object {{
        $targetPath = $_.FullName.Replace($sourceFolder, $appPath)
        if ($_.PSIsContainer) {{
            if (!(Test-Path $targetPath)) {{
                New-Item -ItemType Directory -Path $targetPath -Force | Out-Null
            }}
        }} else {{
            $targetDir = Split-Path $targetPath -Parent
            if (!(Test-Path $targetDir)) {{
                New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
            }}
            Copy-Item -Path $_.FullName -Destination $targetPath -Force
        }}
    }}
}} catch {{
    Write-Host ""Error replacing files: $($_.Exception.Message)""
    exit 1
}}

# Clean up temp files
try {{
    Remove-Item -Path $updateSource -Recurse -Force -ErrorAction SilentlyContinue
    $zipFiles = Get-ChildItem -Path (Split-Path $updateSource -Parent) -Filter '*.zip' -ErrorAction SilentlyContinue
    $zipFiles | Remove-Item -Force -ErrorAction SilentlyContinue
}} catch {{
    # Ignore cleanup errors
}}

# Restart the service if it was installed
try {{
    $service = Get-Service -Name 'LiteTaskService' -ErrorAction SilentlyContinue
    if ($service) {{
        try {{
            Start-Service -Name 'LiteTaskService' -ErrorAction Stop
        }} catch {{
            # Direct start failed (e.g. insufficient privileges), try elevated
            Start-Process powershell -ArgumentList '-NoProfile -Command ""Start-Service LiteTaskService""' -Verb RunAs -WindowStyle Hidden -Wait -ErrorAction Stop
        }}
        $service.Refresh()
        $service.WaitForStatus('Running', (New-TimeSpan -Seconds 30)) | Out-Null
    }}
}} catch {{
    # Service not installed, user declined UAC, or failed to start
}}

# Restart the application
Start-Process -FilePath (Join-Path $appPath 'LiteTask.exe')
"

            File.WriteAllText(updateScriptPath, scriptContent)
            _logger?.LogInfo($"Update script created at: {updateScriptPath}")

            Dim startInfo As New ProcessStartInfo() With {
                .FileName = "powershell.exe",
                .Arguments = $"-NoProfile -ExecutionPolicy Bypass -File ""{updateScriptPath}""",
                .WorkingDirectory = _tempPath,
                .UseShellExecute = True,
                .WindowStyle = ProcessWindowStyle.Hidden
            }

            Process.Start(startInfo)
            _logger?.LogInfo("Update script launched.")
        End Sub

        Private Sub CleanupTempFiles(zipPath As String, extractPath As String)
            Try
                If zipPath IsNot Nothing AndAlso File.Exists(zipPath) Then
                    File.Delete(zipPath)
                End If
            Catch
            End Try
            Try
                If extractPath IsNot Nothing AndAlso Directory.Exists(extractPath) Then
                    Directory.Delete(extractPath, True)
                End If
            Catch
            End Try
        End Sub

        Private Function IsNewerVersion(current As String, remote As String) As Boolean
            Try
                Dim currentParts = current.Split("."c).Select(Function(p) Integer.Parse(p)).ToArray()
                Dim remoteParts = remote.Split("."c).Select(Function(p) Integer.Parse(p)).ToArray()

                Dim maxLength = Math.Max(currentParts.Length, remoteParts.Length)
                For i = 0 To maxLength - 1
                    Dim c = If(i < currentParts.Length, currentParts(i), 0)
                    Dim r = If(i < remoteParts.Length, remoteParts(i), 0)
                    If r > c Then Return True
                    If r < c Then Return False
                Next
                Return False
            Catch
                Return False
            End Try
        End Function

        Public Sub Dispose()
            _httpClient?.Dispose()
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
    End Class
End Namespace
