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
                Dim doc = JsonDocument.Parse(json)
                Dim root = doc.RootElement

                Dim tagName = root.GetProperty("tag_name").GetString()
                Dim remoteVersion = tagName.TrimStart("v"c)
                Dim currentVersion = Application.ProductVersion

                Dim info As New UpdateInfo With {
                    .CurrentVersion = currentVersion,
                    .RemoteVersion = remoteVersion,
                    .IsUpdateAvailable = IsNewerVersion(currentVersion, remoteVersion),
                    .ReleaseNotes = If(root.TryGetProperty("body", Nothing), root.GetProperty("body").GetString(), ""),
                    .ReleaseName = If(root.TryGetProperty("name", Nothing), root.GetProperty("name").GetString(), tagName)
                }

                ' Find the zip asset download URL
                If root.TryGetProperty("assets", Nothing) Then
                    For Each asset In root.GetProperty("assets").EnumerateArray()
                        Dim assetName = asset.GetProperty("name").GetString()
                        If assetName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase) Then
                            info.DownloadUrl = asset.GetProperty("browser_download_url").GetString()
                            info.AssetName = assetName
                            Exit For
                        End If
                    Next
                End If

                _logger?.LogInfo($"Update check complete. Current: {currentVersion}, Remote: {remoteVersion}, Update available: {info.IsUpdateAvailable}")
                Return info

            Catch ex As Exception
                _logger?.LogError($"Error checking for updates: {ex.Message}")
                Throw
            End Try
        End Function

        ''' <summary>
        ''' Downloads, extracts, runs PostBuild, stops service, and launches the update script.
        ''' </summary>
        Public Async Function DownloadAndApplyUpdateAsync(updateInfo As UpdateInfo, progress As IProgress(Of String)) As Task
            Dim zipPath As String = Nothing
            Dim extractPath As String = Nothing

            Try
                Directory.CreateDirectory(_tempPath)
                zipPath = Path.Combine(_tempPath, If(updateInfo.AssetName, "LiteTask-update.zip"))
                extractPath = Path.Combine(_tempPath, "update")

                ' Clean previous update attempt
                If Directory.Exists(extractPath) Then
                    Directory.Delete(extractPath, True)
                End If

                ' Step 1: Download
                progress?.Report("Downloading update...")
                _logger?.LogInfo($"Downloading update from {updateInfo.DownloadUrl}")
                Await DownloadFileAsync(updateInfo.DownloadUrl, zipPath)
                _logger?.LogInfo("Download complete.")

                ' Step 2: Extract
                progress?.Report("Extracting files...")
                _logger?.LogInfo($"Extracting to {extractPath}")
                If Directory.Exists(extractPath) Then
                    Directory.Delete(extractPath, True)
                End If
                ZipFile.ExtractToDirectory(zipPath, extractPath)
                _logger?.LogInfo("Extraction complete.")

                ' Step 3: Run PostBuild script
                progress?.Report("Running post-build script...")
                Dim postBuildScript = FindFile(extractPath, "LiteTask-Post-Build.ps1")
                If postBuildScript IsNot Nothing Then
                    _logger?.LogInfo($"Running PostBuild script: {postBuildScript}")
                    RunPostBuildScript(postBuildScript)
                    _logger?.LogInfo("PostBuild script completed.")
                Else
                    _logger?.LogInfo("No PostBuild script found, skipping.")
                End If

                ' Step 4: Stop service if running
                progress?.Report("Stopping service...")
                StopServiceIfRunning()

                ' Step 5: Launch update script and exit
                progress?.Report("Launching update process...")
                LaunchUpdateScript(extractPath)

                ' Step 6: Exit application
                _logger?.LogInfo("Exiting application for update...")
                Application.Exit()

            Catch ex As Exception
                _logger?.LogError($"Error applying update: {ex.Message}")
                ' Clean up on error
                CleanupTempFiles(zipPath, extractPath)
                Throw
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

        Private Sub RunPostBuildScript(scriptPath As String)
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

                Using proc = Process.Start(startInfo)
                    proc.WaitForExit(60000) ' 60 second timeout
                    If Not proc.HasExited Then
                        proc.Kill()
                        _logger?.LogError("PostBuild script timed out after 60 seconds.")
                    Else
                        Dim output = proc.StandardOutput.ReadToEnd()
                        Dim errors = proc.StandardError.ReadToEnd()
                        If Not String.IsNullOrWhiteSpace(output) Then
                            _logger?.LogInfo($"PostBuild output: {output}")
                        End If
                        If Not String.IsNullOrWhiteSpace(errors) Then
                            _logger?.LogError($"PostBuild errors: {errors}")
                        End If
                    End If
                End Using
            Catch ex As Exception
                _logger?.LogError($"Error running PostBuild script: {ex.Message}")
            End Try
        End Sub

        Private Sub StopServiceIfRunning()
            Try
                Using sc = New ServiceController(ServiceName)
                    If sc.Status = ServiceControllerStatus.Running OrElse
                       sc.Status = ServiceControllerStatus.StartPending Then
                        _logger?.LogInfo($"Stopping service {ServiceName}...")
                        sc.Stop()
                        sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(30))
                        _logger?.LogInfo($"Service {ServiceName} stopped.")
                    Else
                        _logger?.LogInfo($"Service {ServiceName} is not running (status: {sc.Status}).")
                    End If
                End Using
            Catch ex As InvalidOperationException
                ' Service not installed
                _logger?.LogInfo($"Service {ServiceName} is not installed, skipping stop.")
            Catch ex As Exception
                _logger?.LogError($"Error stopping service: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Creates and launches a PowerShell script that waits for the app to exit,
        ''' then replaces the files and restarts the application.
        ''' </summary>
        Private Sub LaunchUpdateScript(extractPath As String)
            Dim updateScriptPath = Path.Combine(_tempPath, "LiteTask-Update.ps1")
            Dim currentPid = Environment.ProcessId

            ' Find the actual content folder inside the extracted path
            ' It could be directly in extractPath or in a subfolder
            Dim scriptContent = $"
# LiteTask Update Script - Auto-generated
# Wait for the application to exit
$processId = {currentPid}
$appPath = '{_appPath.Replace("'", "''")}'
$updateSource = '{extractPath.Replace("'", "''")}'

Write-Host 'Waiting for LiteTask to exit...'
try {{
    $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
    if ($process) {{
        $process.WaitForExit(30000)
        if (!$process.HasExited) {{
            Write-Host 'Force stopping LiteTask...'
            $process | Stop-Process -Force
            Start-Sleep -Seconds 2
        }}
    }}
}} catch {{
    # Process already exited
}}

Start-Sleep -Seconds 2

Write-Host 'Replacing files...'

# Determine the source folder (handle nested folder structure)
$sourceFolder = $updateSource
$subDirs = Get-ChildItem -Path $updateSource -Directory
if ($subDirs.Count -eq 1 -and (Test-Path (Join-Path $subDirs[0].FullName 'LiteTask.exe'))) {{
    $sourceFolder = $subDirs[0].FullName
}} elseif (!(Test-Path (Join-Path $updateSource 'LiteTask.exe'))) {{
    # Search deeper for LiteTask.exe
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
            Write-Host ""Updated: $targetPath""
        }}
    }}
    Write-Host 'File replacement complete.'
}} catch {{
    Write-Host ""Error replacing files: $($_.Exception.Message)""
    Read-Host 'Press Enter to exit'
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

# Restart the application
Write-Host 'Restarting LiteTask...'
Start-Process -FilePath (Join-Path $appPath 'LiteTask.exe')
Write-Host 'Update complete!'
"

            File.WriteAllText(updateScriptPath, scriptContent)
            _logger?.LogInfo($"Update script created at: {updateScriptPath}")

            Dim startInfo As New ProcessStartInfo() With {
                .FileName = "powershell.exe",
                .Arguments = $"-NoProfile -ExecutionPolicy Bypass -File ""{updateScriptPath}""",
                .WorkingDirectory = _tempPath,
                .UseShellExecute = True,
                .CreateNoWindow = False
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
