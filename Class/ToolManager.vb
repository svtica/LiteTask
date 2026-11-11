Imports System.Collections.Concurrent
Imports System.IO.Compression
Imports System.Net.Http
Imports System.Text.Json

Namespace LiteTask

    ''' <summary>Outcome of a tool download/installation.</summary>
    Public Class ToolInstallResult
        Public Property Tool As ToolDefinition
        Public Property Success As Boolean
        Public Property Version As String
        Public Property Message As String
        Public Property Signature As AuthenticodeResult
    End Class

    Public Class ToolManager
        Public ReadOnly _toolsPath As String
        Private ReadOnly _tools As Dictionary(Of String, String)
        Private ReadOnly _httpClient As HttpClient
        Private _currentExecutionTool As String = "PsExec64.exe"
        Private ReadOnly _xmlManager As XMLManager
        Private ReadOnly _logger As Logger
        Private ReadOnly _tempPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LiteTaskData", "temp")
        Private ReadOnly _installLocks As New ConcurrentDictionary(Of String, SemaphoreSlim)(StringComparer.OrdinalIgnoreCase)
        Private Const VERSION_FILE_SUFFIX As String = ".version"

        Public ReadOnly Property ToolsPath As String
            Get
                Return _toolsPath
            End Get
        End Property

        Public Property CurrentExecutionTool As String
            Get
                Return _currentExecutionTool
            End Get
            Set(value As String)
                If _tools.ContainsKey(value) Then
                    _currentExecutionTool = value
                    ' Save the setting when changed
                    _xmlManager.SaveToolSettings(value)
                    _logger?.LogInfo($"Execution tool changed to: {value}")
                Else
                    Throw New ArgumentException($"Tool '{value}' is not available.")
                End If
            End Set
        End Property

        ''' <summary>Optional tools that can be downloaded from Tools &gt; Manage Tools.</summary>
        Public ReadOnly Property Catalog As IReadOnlyList(Of ToolDefinition)
            Get
                Return ToolCatalog.All
            End Get
        End Property

        Public Sub New(toolsPath As String)
            Try
                _toolsPath = toolsPath
                _tools = New Dictionary(Of String, String)
                _xmlManager = ApplicationContainer.GetService(Of XMLManager)()
                _logger = ApplicationContainer.GetService(Of Logger)()

                ' Execution tools: PsExec64 is embedded, sqlcmd is downloaded on demand
                AddToolSafely("PsExec64.exe", "embedded")
                AddToolSafely("sqlcmd.exe", ToolCatalog.SqlCmdId)

                _httpClient = New HttpClient() With {.Timeout = TimeSpan.FromMinutes(10)}
                ' GitHub's REST API rejects requests without a User-Agent
                _httpClient.DefaultRequestHeaders.Add("User-Agent", "LiteTask-ToolManager")
                LoadToolSetting()

                ' Tools copied by hand from a browser download carry the Mark of the Web
                Dim unblocked = MarkOfTheWeb.RemoveFromDirectory(_toolsPath)
                If unblocked > 0 Then _logger?.LogInfo($"Removed the Mark of the Web from {unblocked} file(s) in {_toolsPath}")
            Catch ex As Exception
                _logger?.LogError($"Error initializing ToolManager: {ex.Message}")
                Throw
            End Try
        End Sub

        Private Sub AddToolSafely(key As String, value As String)
            If Not _tools.ContainsKey(key) Then
                _tools.Add(key, value)
            End If
        End Sub

        ''' <summary>Presence of each tool, for Tools &gt; Check Tools.</summary>
        Public Function DetectTools() As Dictionary(Of String, Boolean)
            Dim result As New Dictionary(Of String, Boolean) From {
                {"PsExec64.exe", True} ' embedded in LiteTask.exe, extracted on first use
            }
            For Each tool In Catalog
                result(tool.ExecutableName) = IsInstalled(tool)
            Next
            Return result
        End Function

        Public Sub Dispose()
            _httpClient.Dispose()
        End Sub

#Region "Installed tools"

        ''' <summary>
        ''' Full path of the installed executable, or Nothing. Looks in the tool's install
        ''' folder first, then at the tools root where versions up to 1.2.2 shipped it.
        ''' </summary>
        Public Function GetInstalledExecutable(tool As ToolDefinition) As String
            Dim candidates As New List(Of String)
            If Not String.IsNullOrEmpty(tool.InstallFolder) Then
                candidates.Add(Path.Combine(_toolsPath, tool.InstallFolder, tool.ExecutableName))
            End If
            candidates.Add(Path.Combine(_toolsPath, tool.ExecutableName))
            Return candidates.FirstOrDefault(Function(p) File.Exists(p))
        End Function

        Public Function IsInstalled(tool As ToolDefinition) As Boolean
            Return GetInstalledExecutable(tool) IsNot Nothing
        End Function

        ''' <summary>Release tag recorded at installation, "" when unknown (shipped with an older package).</summary>
        Public Function GetInstalledVersion(tool As ToolDefinition) As String
            Try
                Dim versionFile = Path.Combine(_toolsPath, tool.Id & VERSION_FILE_SUFFIX)
                If File.Exists(versionFile) Then Return File.ReadAllText(versionFile).Trim()
            Catch
            End Try
            Return ""
        End Function

        ''' <summary>Tools chosen in Tools &gt; Manage Tools (ToolSettings/SelectedTools).</summary>
        Public Function GetSelectedToolIds() As HashSet(Of String)
            Dim stored = _xmlManager.ReadValue("ToolSettings", "SelectedTools", Nothing)
            If stored Is Nothing Then
                Return ToolCatalog.DefaultSelection(AddressOf IsInstalled)
            End If
            Return ToolCatalog.ParseSelection(stored)
        End Function

        Public Sub SaveSelectedToolIds(ids As IEnumerable(Of String))
            _xmlManager.WriteValue("ToolSettings", "SelectedTools", ToolCatalog.FormatSelection(ids))
        End Sub

#End Region

#Region "Download / install / remove"

        ''' <summary>
        ''' Downloads the tool's release asset from GitHub and installs it under
        ''' LiteTaskData\tools. Replaces an existing installation.
        ''' </summary>
        Public Async Function InstallToolAsync(tool As ToolDefinition, Optional progress As Action(Of String) = Nothing) As Task(Of ToolInstallResult)
            Dim result As New ToolInstallResult With {.Tool = tool}
            Dim installLock = _installLocks.GetOrAdd(tool.Id, Function(k) New SemaphoreSlim(1, 1))
            Await installLock.WaitAsync()
            Dim workDir = Path.Combine(_tempPath, $"tool_{tool.Id}_{Guid.NewGuid():N}")
            Try
                Report(progress, String.Format(Tr("Tools.Progress.Querying", "{0}: querying {1} releases..."), tool.DisplayName, tool.Repository))

                ' 1. Release metadata
                Dim tag As String
                Dim assets As New List(Of KeyValuePair(Of String, String))
                Using response = Await _httpClient.GetAsync(ToolCatalog.GetReleaseApiUrl(tool))
                    If Not response.IsSuccessStatusCode Then
                        Throw New InvalidOperationException($"GitHub returned {CInt(response.StatusCode)} {response.ReasonPhrase} for {tool.Repository}")
                    End If
                    Using doc = JsonDocument.Parse(Await response.Content.ReadAsStringAsync())
                        tag = doc.RootElement.GetProperty("tag_name").GetString()
                        Dim assetsElement As JsonElement
                        If doc.RootElement.TryGetProperty("assets", assetsElement) Then
                            For Each asset In assetsElement.EnumerateArray()
                                assets.Add(New KeyValuePair(Of String, String)(
                                    asset.GetProperty("name").GetString(),
                                    asset.GetProperty("browser_download_url").GetString()))
                            Next
                        End If
                    End Using
                End Using

                Dim selected = ToolCatalog.SelectAsset(tool, assets)
                If Not selected.HasValue Then
                    Throw New InvalidOperationException($"Release {tag} of {tool.Repository} has no downloadable asset matching '{tool.AssetPattern}'")
                End If
                Dim assetName = Path.GetFileName(selected.Value.Key)

                ' 2. Download
                Directory.CreateDirectory(workDir)
                Dim downloadPath = Path.Combine(workDir, assetName)
                Report(progress, String.Format(Tr("Tools.Progress.Downloading", "{0}: downloading {1} ({2})..."), tool.DisplayName, assetName, tag))
                Using response = Await _httpClient.GetAsync(selected.Value.Value, HttpCompletionOption.ResponseHeadersRead)
                    response.EnsureSuccessStatusCode()
                    Using content = Await response.Content.ReadAsStreamAsync(),
                          fileStream As New FileStream(downloadPath, FileMode.Create, FileAccess.Write, FileShare.None)
                        Await content.CopyToAsync(fileStream)
                    End Using
                End Using

                ' 3. Locate the executable (zip: extract; ExtractToDirectory rejects entries escaping the folder)
                Dim sourceDir As String
                If assetName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase) Then
                    Dim extractDir = Path.Combine(workDir, "extract")
                    ZipFile.ExtractToDirectory(downloadPath, extractDir)
                    Dim exe = Directory.GetFiles(extractDir, tool.ExecutableName, SearchOption.AllDirectories).
                        OrderBy(Function(p) p.Length).FirstOrDefault()
                    If exe Is Nothing Then
                        Throw New FileNotFoundException($"{tool.ExecutableName} not found in {assetName}")
                    End If
                    sourceDir = Path.GetDirectoryName(exe)
                Else
                    sourceDir = Path.Combine(workDir, "single")
                    Directory.CreateDirectory(sourceDir)
                    File.Move(downloadPath, Path.Combine(sourceDir, tool.ExecutableName))
                End If

                ' 4. Signature: a signature that does not match the file means tampering or corruption
                Dim signature = AuthenticodeVerifier.Verify(Path.Combine(sourceDir, tool.ExecutableName))
                result.Signature = signature
                _logger?.LogInfo($"Tool {tool.Id} {tag} signature: {signature}")
                If signature.Status = AuthenticodeStatus.HashMismatch Then
                    Throw New InvalidOperationException($"{tool.ExecutableName} signature does not match its content; installation refused")
                End If
                If Not signature.IsSignatureValid Then
                    Report(progress, String.Format(Tr("Tools.Progress.NotSigned", "{0}: warning, {1} is not signed ({2})"), tool.DisplayName, tool.ExecutableName, signature.Status))
                End If

                ' 5. Install
                Directory.CreateDirectory(_toolsPath)
                If String.IsNullOrEmpty(tool.InstallFolder) Then
                    File.Copy(Path.Combine(sourceDir, tool.ExecutableName), Path.Combine(_toolsPath, tool.ExecutableName), True)
                Else
                    Dim target = Path.Combine(_toolsPath, tool.InstallFolder)
                    If Directory.Exists(target) Then Directory.Delete(target, True)
                    CopyDirectory(sourceDir, target)
                    ' The copy shipped at the tools root by older packages would shadow nothing
                    ' (the install folder is looked up first) but would never be updated.
                    TryDeleteFile(Path.Combine(_toolsPath, tool.ExecutableName))
                    TryDeleteFile(Path.Combine(_toolsPath, tool.ExecutableName & ".config"))
                End If
                File.WriteAllText(Path.Combine(_toolsPath, tool.Id & VERSION_FILE_SUFFIX), tag)
                If String.IsNullOrEmpty(tool.InstallFolder) Then
                    MarkOfTheWeb.RemoveFromFile(Path.Combine(_toolsPath, tool.ExecutableName))
                Else
                    MarkOfTheWeb.RemoveFromDirectory(Path.Combine(_toolsPath, tool.InstallFolder))
                End If

                result.Success = True
                result.Version = tag
                result.Message = String.Format(Tr("Tools.Progress.Installed", "{0} {1} installed"), tool.DisplayName, tag)
                _logger?.LogInfo($"Tool {tool.Id} {tag} installed from {selected.Value.Value}")
                Report(progress, result.Message)
            Catch ex As Exception
                result.Success = False
                result.Message = $"{tool.DisplayName}: {ex.Message}"
                _logger?.LogError($"Error installing tool {tool.Id}: {ex.Message}")
                Report(progress, result.Message)
            Finally
                Try
                    If Directory.Exists(workDir) Then Directory.Delete(workDir, True)
                Catch
                End Try
                installLock.Release()
            End Try
            Return result
        End Function

        ''' <summary>
        ''' Makes sure a tool is present, installing it when missing (used when a task
        ''' needs sqlcmd on a fresh installation). Returns the executable path or Nothing.
        ''' </summary>
        Public Async Function EnsureInstalledAsync(tool As ToolDefinition) As Task(Of String)
            Dim existing = GetInstalledExecutable(tool)
            If existing IsNot Nothing Then Return existing
            _logger?.LogWarning($"{tool.ExecutableName} is missing; downloading it from {tool.Repository}")
            Dim installResult = Await InstallToolAsync(tool)
            Return If(installResult.Success, GetInstalledExecutable(tool), Nothing)
        End Function

        ''' <summary>Deletes an installed tool (install folder, legacy root copy, version file).</summary>
        Public Function RemoveTool(tool As ToolDefinition) As Boolean
            Try
                If Not String.IsNullOrEmpty(tool.InstallFolder) Then
                    Dim folder = Path.Combine(_toolsPath, tool.InstallFolder)
                    If Directory.Exists(folder) Then Directory.Delete(folder, True)
                End If
                Dim rootExe = Path.Combine(_toolsPath, tool.ExecutableName)
                If File.Exists(rootExe) Then File.Delete(rootExe)
                TryDeleteFile(rootExe & ".config")
                TryDeleteFile(Path.Combine(_toolsPath, tool.Id & VERSION_FILE_SUFFIX))
                _logger?.LogInfo($"Tool {tool.Id} removed")
                Return True
            Catch ex As Exception
                _logger?.LogError($"Error removing tool {tool.Id}: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>Downloads (or updates) every selected tool.</summary>
        Public Async Function DownloadAndUpdateAllToolsAsync(Optional progress As Action(Of String) = Nothing) As Task(Of List(Of ToolInstallResult))
            Dim results As New List(Of ToolInstallResult)
            Dim selected = GetSelectedToolIds()
            For Each tool In Catalog.Where(Function(t) selected.Contains(t.Id))
                results.Add(Await InstallToolAsync(tool, progress))
            Next
            Return results
        End Function

        Private Shared Sub CopyDirectory(sourceDir As String, targetDir As String)
            Directory.CreateDirectory(targetDir)
            For Each subDir In Directory.GetDirectories(sourceDir, "*", SearchOption.AllDirectories)
                Directory.CreateDirectory(Path.Combine(targetDir, Path.GetRelativePath(sourceDir, subDir)))
            Next
            For Each sourceFile In Directory.GetFiles(sourceDir, "*", SearchOption.AllDirectories)
                File.Copy(sourceFile, Path.Combine(targetDir, Path.GetRelativePath(sourceDir, sourceFile)), True)
            Next
        End Sub

        Private Shared Sub TryDeleteFile(filePath As String)
            Try
                If File.Exists(filePath) Then File.Delete(filePath)
            Catch
            End Try
        End Sub

        ''' <summary>UI text only (the Manage Tools window); log lines stay in English.</summary>
        Private Shared Function Tr(key As String, defaultValue As String) As String
            Try
                Return TranslationManager.Instance.GetTranslation(key, defaultValue)
            Catch
                Return defaultValue
            End Try
        End Function

        Private Shared Sub Report(progress As Action(Of String), message As String)
            Try
                progress?.Invoke(message)
            Catch
            End Try
        End Sub

#End Region

        Public Function GetCurrentToolPath() As String
            Return GetToolPath(_currentExecutionTool)
        End Function

        Public Function GetToolPath(toolName As String) As String
            Return Path.Combine(_toolsPath, toolName)
        End Function

        ''' <summary>Starts LitePM filtered on LiteTask processes. Returns False when LitePM is not installed.</summary>
        Public Function LaunchLitePM(processName As String) As Boolean
            Dim litePmPath = GetInstalledExecutable(ToolCatalog.Find(ToolCatalog.LitePMId))
            If litePmPath Is Nothing Then Return False

            Dim startInfo = New ProcessStartInfo() With {
                .FileName = litePmPath,
                .WorkingDirectory = Path.GetDirectoryName(litePmPath),
                .Arguments = "-filter ""Lite""",
                .UseShellExecute = True
            }

            ' Started through the shell: a marked LitePM would show the security warning
            MarkOfTheWeb.RemoveFromDirectory(Path.GetDirectoryName(litePmPath))
            Dim proc = Process.Start(startInfo)
            proc?.Dispose()
            Return True
        End Function

        Private Sub LoadToolSetting()
            Try
                Dim settings = _xmlManager.GetToolSettings()
                Dim savedTool = settings("ExecutionTool")

                ' Verify the saved tool exists before using it
                If _tools.ContainsKey(savedTool) Then
                    _currentExecutionTool = savedTool
                Else
                    _currentExecutionTool = "PsExec64.exe"  ' Default if saved tool not found
                    _xmlManager.SaveToolSettings(_currentExecutionTool)
                End If

                _logger?.LogInfo($"Loaded execution tool setting: {_currentExecutionTool}")
            Catch ex As Exception
                _logger?.LogError($"Error loading tool setting: {ex.Message}")
                _currentExecutionTool = "PsExec64.exe"  ' Fallback to default
            End Try
        End Sub

    End Class
End Namespace
