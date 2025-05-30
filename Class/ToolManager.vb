Imports System.IO.Compression
Imports System.Net.Http

Namespace LiteTask
    Public Class ToolManager
        Public ReadOnly _toolsPath As String
        Private ReadOnly _tools As Dictionary(Of String, String)
        Private ReadOnly _httpClient As HttpClient
        Private _currentExecutionTool As String = "PsExec64.exe"
        Private ReadOnly _xmlManager As XMLManager
        Private ReadOnly _logger As Logger
        Private Const SqlcmdUrl As String = "https://github.com/microsoft/go-sqlcmd/releases/download/v1.8.2/sqlcmd-windows-amd64.zip"
        Private ReadOnly _tempPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LiteTaskData", "temp")

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

        Public Sub New(toolsPath As String)
            Try
                _toolsPath = toolsPath
                _tools = New Dictionary(Of String, String)
                _xmlManager = ApplicationContainer.GetService(Of XMLManager)()
                _logger = ApplicationContainer.GetService(Of Logger)()

                ' Add embedded PsExec64 tool
                AddToolSafely("PsExec64.exe", "embedded")
                AddToolSafely("sqlcmd.exe", SqlcmdUrl)

                _httpClient = New HttpClient()
                LoadToolSetting()

                '_logger?.LogInfo($"ToolManager initialized with execution tool: {_currentExecutionTool}")
            Catch ex As Exception
                _logger?.LogError($"Error initializing ToolManager: {ex.Message}")
                Throw
            End Try
        End Sub

        Private Sub AddToolSafely(key As String, value As String)
            If Not _tools.ContainsKey(key) Then
                _tools.Add(key, value)
            Else
                ' Optionally log a warning about duplicate key
                ' _logger.LogWarning($"Duplicate tool key found: {key}")
            End If
        End Sub

        Public Function DetectTools() As Dictionary(Of String, Boolean)
            Dim result As New Dictionary(Of String, Boolean)
            For Each tool In _tools.Keys
                result(tool) = File.Exists(Path.Combine(_toolsPath, tool))
            Next
            Return result
        End Function

        Public Sub Dispose()
            _httpClient.Dispose()
        End Sub

        Public Async Function DownloadAndUpdateToolAsync(toolName As String, url As String) As Task(Of Boolean)
            Try
                Dim toolPath = Path.Combine(_toolsPath, toolName)
                Using response As HttpResponseMessage = Await _httpClient.GetAsync(url)
                    If response.IsSuccessStatusCode Then
                        Using stream As Stream = Await response.Content.ReadAsStreamAsync()
                            Using fileStream As New FileStream(toolPath, FileMode.Create, FileAccess.Write, FileShare.None)
                                Await stream.CopyToAsync(fileStream)
                            End Using
                        End Using
                        Return True
                    Else
                        Console.WriteLine($"Error downloading {toolName}: HTTP status code {response.StatusCode}")
                        Return False
                    End If
                End Using
            Catch ex As Exception
                Console.WriteLine($"Error downloading {toolName}: {ex.Message}")
                Return False
            End Try
        End Function

        Public Async Function DownloadAndInstallSqlcmdAsync() As Task(Of Boolean)
            Try
                ' Ensure temp directory exists
                Directory.CreateDirectory(_tempPath)

                Dim zipFileName = "sqlcmd-windows-amd64.zip"
                Dim zipPath = Path.Combine(_tempPath, zipFileName)

                ' Download the zip file
                Await DownloadFileAsync(SqlcmdUrl, zipPath)

                ' Extract the zip file
                Dim extractPath = Path.Combine(_tempPath, "sqlcmd_extract")
                ExtractZipFile(zipPath, extractPath)

                ' Copy sqlcmd.exe to the tools folder
                Dim sourceFile = Path.Combine(extractPath, "sqlcmd.exe")
                Dim destFile = Path.Combine(_toolsPath, "sqlcmd.exe")
                File.Copy(sourceFile, destFile, True)

                ' Clean up
                File.Delete(zipPath)
                Directory.Delete(extractPath, True)

                ' Add sqlcmd.exe to the tools dictionary
                AddToolSafely("sqlcmd.exe", SqlcmdUrl)

                _logger?.LogInfo("sqlcmd.exe has been successfully installed.")
                Return True
            Catch ex As Exception
                _logger?.LogError($"Error installing sqlcmd.exe: {ex.Message}")
                Return False
            End Try
        End Function


        Public Async Function DownloadAndUpdateAllToolsAsync() As Task(Of Dictionary(Of String, Boolean))
            Dim result As New Dictionary(Of String, Boolean)
            For Each kvp In _tools
                result(kvp.Key) = Await DownloadAndUpdateToolAsync(kvp.Key, kvp.Value)
            Next
            Return result
        End Function

        Private Async Function DownloadFileAsync(url As String, filePath As String) As Task
            Using response As HttpResponseMessage = Await _httpClient.GetAsync(url)
                response.EnsureSuccessStatusCode()
                Using content = Await response.Content.ReadAsStreamAsync()
                    Using fileStream As New FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None)
                        Await content.CopyToAsync(fileStream)
                    End Using
                End Using
            End Using
        End Function

        Private Sub ExtractZipFile(zipPath As String, extractPath As String)
            ZipFile.ExtractToDirectory(zipPath, extractPath)
        End Sub

        Public Function GetCurrentToolPath() As String
            Return GetToolPath(_currentExecutionTool)
        End Function

        Public Function GetToolPath(toolName As String) As String
            Return Path.Combine(_toolsPath, toolName)
        End Function

        Public Sub LaunchLitePM(processName As String)
            Dim procExpPath = Path.Combine(_toolsPath, "LitePM.exe")

            ' Launch Process Explorer with a filter for LiteTask processes
            Dim startInfo = New ProcessStartInfo() With {
            .FileName = procExpPath,
            .Arguments = "-filter ""Lite""",
            .UseShellExecute = True
        }

            Process.Start(startInfo)
        End Sub

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