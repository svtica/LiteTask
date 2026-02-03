Imports System.Xml
Imports System.IO
Imports LiteTask.LiteTask.ScheduledTask

Namespace LiteTask
    Public Class XMLManager
        Private _filePath As String
        Private _logger As Logger
        Private ReadOnly _xmlLock As New ReaderWriterLockSlim()
        Private ReadOnly _backupPath As String
        Private ReadOnly _tempPath As String
        
        ' Simple cache for frequently accessed config values
        Private _configCache As New Dictionary(Of String, String)
        Private _cacheExpiry As DateTime = DateTime.MinValue
        Private ReadOnly _cacheTimeout As TimeSpan = TimeSpan.FromMinutes(5)

        Public Sub New(filePath As String)
            If String.IsNullOrWhiteSpace(filePath) Then
                Throw New ArgumentException("File path cannot be null or empty.", NameOf(filePath))
            End If

            _filePath = filePath
            _backupPath = Path.Combine(Path.GetDirectoryName(filePath), "backup")
            _tempPath = Path.Combine(Application.StartupPath, "LiteTaskData", "temp")

            ' Create directories with proper error handling
            Try
                Directory.CreateDirectory(Path.GetDirectoryName(filePath))
                Directory.CreateDirectory(_backupPath)
                Directory.CreateDirectory(_tempPath)
            Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is UnauthorizedAccessException
                Throw New InvalidOperationException($"Unable to create required directories: {ex.Message}", ex)
            End Try

            ' Ensure config file exists with proper locking
            EnsureConfigFileExists()
        End Sub

        Private Sub AddElement(doc As XmlDocument, parent As XmlElement, name As String, value As String)
            If Not String.IsNullOrEmpty(value) Then
                Dim element As XmlElement = doc.CreateElement(name)
                element.InnerText = SecurityElement.Escape(value)
                parent.AppendChild(element)
            End If
        End Sub

        Private Sub CreateBackup()
            Try
                If Not File.Exists(_filePath) Then Return

                Dim backupFileName = $"backup_{DateTime.Now:yyyyMMddHHmmss}.xml"
                Dim backupFilePath = Path.Combine(_backupPath, backupFileName)

                File.Copy(_filePath, backupFilePath, True)

            Catch ex As Exception
                _logger?.LogError($"Error creating backup: {ex.Message}")
            End Try
        End Sub

        Private Sub CreateDefaultConfig()
            Try
                Dim xmlDoc As New XmlDocument()
                Dim declaration = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", Nothing)
                xmlDoc.AppendChild(declaration)

                Dim root = xmlDoc.CreateElement("LiteTaskSettings")
                xmlDoc.AppendChild(root)

                ' Tool Settings
                Dim toolSection = xmlDoc.CreateElement("ToolSettings")
                Dim execToolElem = xmlDoc.CreateElement("ExecutionTool")
                execToolElem.InnerText = "PsExec64.exe"
                toolSection.AppendChild(execToolElem)
                root.AppendChild(toolSection)

                ' SQL Configuration
                Dim sqlSection = xmlDoc.CreateElement("SqlConfiguration")
                sqlSection.AppendChild(CreateElement(xmlDoc, "DefaultServer", ""))
                sqlSection.AppendChild(CreateElement(xmlDoc, "DefaultDatabase", ""))
                sqlSection.AppendChild(CreateElement(xmlDoc, "CommandTimeout", "300"))
                sqlSection.AppendChild(CreateElement(xmlDoc, "MaxBatchSize", "1000"))
                root.AppendChild(sqlSection)

                ' Tasks section
                Dim tasksSection = xmlDoc.CreateElement("Tasks")
                root.AppendChild(tasksSection)

                xmlDoc.Save(_filePath)
            Catch ex As Exception
                _logger?.LogError($"Error creating default configuration: {ex.Message}")
                Throw
            End Try
        End Sub

        Private Function CreateElement(xmlDoc As XmlDocument, name As String, value As String) As XmlElement
            Dim element = xmlDoc.CreateElement(name)
            element.InnerText = value
            Return element
        End Function

        'Private Sub CreateLiteRunDefaultsSection()
        '    Try
        '        Dim xmlDoc As New XmlDocument()
        '        xmlDoc.Load(_filePath)

        '        Dim root = xmlDoc.SelectSingleNode("LiteTaskSettings")
        '        If root Is Nothing Then
        '            root = xmlDoc.CreateElement("LiteTaskSettings")
        '            xmlDoc.AppendChild(root)
        '        End If

        '        Dim liteRunSection = xmlDoc.CreateElement("LiteRunDefaults")
        '        liteRunSection.AppendChild(CreateElement(xmlDoc, "Timeout", "300"))
        '        liteRunSection.AppendChild(CreateElement(xmlDoc, "LogOutputPath", Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LiteRunLogs")))

        '        root.AppendChild(liteRunSection)
        '        xmlDoc.Save(_filePath)
        '    Catch ex As Exception
        '        _logger?.LogError($"Error creating LiteRunDefaults section: {ex.Message}")
        '    End Try
        'End Sub

        Private Function CreateTaskElement(xmlDoc As XmlDocument, task As ScheduledTask) As XmlElement
            Try
                Dim taskElement As XmlElement = xmlDoc.CreateElement("Task")
                taskElement.SetAttribute("Name", SecurityElement.Escape(task.Name))

                AddElement(xmlDoc, taskElement, "Description", task.Description)
                AddElement(xmlDoc, taskElement, "StartTime", task.StartTime.ToString("o"))
                AddElement(xmlDoc, taskElement, "RecurrenceType", CType(task.Schedule, Integer).ToString())
                AddElement(xmlDoc, taskElement, "Interval", task.Interval.TotalMinutes.ToString())
                AddElement(xmlDoc, taskElement, "DailyTimes", String.Join(",", task.DailyTimes.Select(Function(t) t.ToString())))
                AddElement(xmlDoc, taskElement, "MonthlyDay", task.MonthlyDay.ToString())
                AddElement(xmlDoc, taskElement, "MonthlyTime", task.MonthlyTime.ToString())
                AddElement(xmlDoc, taskElement, "NextRunTime", task.NextRunTime.ToString("o"))
                AddElement(xmlDoc, taskElement, "CredentialTarget", task.CredentialTarget)
                AddElement(xmlDoc, taskElement, "AccountType", task.AccountType)
                AddElement(xmlDoc, taskElement, "UserSid", task.UserSid)
                AddElement(xmlDoc, taskElement, "ExecutionMode", CType(task.ExecutionMode, Integer).ToString())

                If task.NextTaskId.HasValue Then
                    AddElement(xmlDoc, taskElement, "NextTaskId", task.NextTaskId.Value.ToString())
                End If

                ' Add actions
                Dim actionsElement = xmlDoc.CreateElement("Actions")
                For Each action In task.Actions
                    Dim actionElement = xmlDoc.CreateElement("Action")
                    AddElement(xmlDoc, actionElement, "Name", action.Name)
                    AddElement(xmlDoc, actionElement, "Order", action.Order.ToString())
                    AddElement(xmlDoc, actionElement, "Type", CType(action.Type, Integer).ToString())
                    AddElement(xmlDoc, actionElement, "Target", action.Target)
                    AddElement(xmlDoc, actionElement, "Parameters", action.Parameters)
                    AddElement(xmlDoc, actionElement, "RequiresElevation", action.RequiresElevation.ToString())

                    If action.DependsOn IsNot Nothing Then
                        AddElement(xmlDoc, actionElement, "DependsOn", action.DependsOn)
                    End If

                    AddElement(xmlDoc, actionElement, "WaitForCompletion", action.WaitForCompletion.ToString())
                    AddElement(xmlDoc, actionElement, "TimeoutMinutes", action.TimeoutMinutes.ToString())
                    AddElement(xmlDoc, actionElement, "RetryCount", action.RetryCount.ToString())
                    AddElement(xmlDoc, actionElement, "RetryDelayMinutes", action.RetryDelayMinutes.ToString())
                    AddElement(xmlDoc, actionElement, "ContinueOnError", action.ContinueOnError.ToString())
                    actionsElement.AppendChild(actionElement)
                Next

                taskElement.AppendChild(actionsElement)
                Return taskElement

            Catch ex As Exception
                _logger?.LogError($"Error creating task element: {ex.Message}")
                Throw
            End Try
        End Function

        Public Sub DeleteTask(taskName As String)
            Try
                _logger?.LogInfo($"Attempting to delete task: {taskName}")

                If String.IsNullOrEmpty(taskName) Then
                    Throw New ArgumentException("Task name cannot be null or empty", NameOf(taskName))
                End If

                ' Load XML document
                Dim xmlDoc As New XmlDocument()
                If Not File.Exists(_filePath) Then
                    _logger?.LogWarning($"XML file not found at {_filePath}")
                    Return
                End If

                xmlDoc.Load(_filePath)

                ' Find the Tasks node, create if it doesn't exist
                Dim tasksNode As XmlNode = xmlDoc.SelectSingleNode("LiteTaskSettings/Tasks")
                If tasksNode Is Nothing Then
                    Dim rootNode As XmlNode = xmlDoc.SelectSingleNode("LiteTaskSettings")
                    If rootNode Is Nothing Then
                        rootNode = xmlDoc.CreateElement("LiteTaskSettings")
                        xmlDoc.AppendChild(rootNode)
                    End If
                    tasksNode = rootNode.AppendChild(xmlDoc.CreateElement("Tasks"))
                End If

                ' Find the task node
                Dim taskNode As XmlNode = tasksNode.SelectSingleNode($"Task[@Name='{taskName}']")
                If taskNode IsNot Nothing Then

                    ' Remove the task node
                    tasksNode.RemoveChild(taskNode)
                    
                    ' Create a backup before saving (using existing backup mechanism)
                    CreateBackup()
                    
                    ' Save the document
                    xmlDoc.Save(_filePath)
                    _logger?.LogInfo($"Task {taskName} deleted successfully from XML")
                Else
                    _logger?.LogWarning($"Task {taskName} not found in XML file")
                End If

            Catch ex As Exception
                _logger?.LogError($"Error deleting task {taskName} from XML: {ex.Message}")
                _logger?.LogError($"StackTrace: {ex.StackTrace}")
                Throw New Exception($"Failed to delete task from XML: {ex.Message}", ex)
            End Try
        End Sub

        Private Sub EnsureConfigFileExists()
            Try
                _xmlLock.EnterWriteLock()

                ' Double-check pattern for thread safety
                If Not File.Exists(_filePath) Then
                    Try
                        CreateDefaultConfig()
                    Catch ex As IOException When File.Exists(_filePath)
                        ' Another thread created the file, we can ignore this
                        Return
                    End Try
                End If
            Finally
                _xmlLock.ExitWriteLock()
            End Try
        End Sub

        Protected Overrides Sub Finalize()
            _xmlLock?.Dispose()
            MyBase.Finalize()
        End Sub

        Public Function GetAllTaskNames() As List(Of String)
            Try
                _logger?.LogInfo("Getting all task names")
                Dim taskNames As New List(Of String)

                If String.IsNullOrEmpty(_filePath) Then
                    _logger?.LogError("File path is null or empty")
                    Return taskNames
                End If

                If Not File.Exists(_filePath) Then
                    _logger?.LogError($"XML file does not exist: {_filePath}")
                    Return taskNames
                End If

                Dim xmlDoc As New XmlDocument()
                xmlDoc.Load(_filePath)

                Dim tasksNode As XmlNode = xmlDoc.SelectSingleNode("LiteTaskSettings/Tasks")
                If tasksNode IsNot Nothing Then
                    Dim taskNodes As XmlNodeList = tasksNode.SelectNodes("Task")
                    For Each taskNode As XmlNode In taskNodes
                        Dim nameAttribute As XmlAttribute = taskNode.Attributes("Name")
                        If nameAttribute IsNot Nothing Then
                            taskNames.Add(nameAttribute.Value)
                        End If
                    Next
                Else
                    _logger?.LogInfo("No 'Tasks' node found in the XML file")
                End If

                _logger?.LogInfo($"Found {taskNames.Count} tasks")
                Return taskNames
            Catch ex As Exception
                _logger?.LogError($"Error getting all task names: {ex.Message}")
                _logger?.LogError($"StackTrace: {ex.StackTrace}")
                Return New List(Of String)() ' Return an empty list instead of throwing
            End Try
        End Function

        Private Function GetDefaultEmailSettings() As Dictionary(Of String, String)
            Return New Dictionary(Of String, String) From {
            {"NotificationsEnabled", "False"},
            {"SmtpServer", ""},
            {"SmtpPort", "25"},
            {"UseSSL", "True"},
            {"EmailFrom", ""},
            {"EmailTo", ""},
            {"UseCredentials", "False"},
            {"CredentialTarget", ""},
            {"AlertLevel", "Error"}
        }
        End Function

        Private Function GetDefaultLogSettings() As Dictionary(Of String, String)
            Return New Dictionary(Of String, String) From {
            {"LogFolder", Path.Combine(Application.StartupPath, "LiteTaskData", "logs")},
            {"LogLevel", "Info"},
            {"MaxLogSize", "10"},
            {"LogRetentionDays", "30"},
            {"AlertLevel", "Error"}
        }
        End Function

        Private Function GetElementValue(node As XmlNode, elementName As String, Optional defaultValue As String = "") As String
            Try
                Dim element As XmlNode = node.SelectSingleNode(elementName)
                Return If(element?.InnerText, defaultValue)
            Catch ex As Exception
                _logger?.LogError($"Error getting element value for {elementName}: {ex.Message}")
                Return defaultValue
            End Try
        End Function

        Public Function GetEmailSettings() As Dictionary(Of String, String)
            Try
                Dim settings As New Dictionary(Of String, String)
                Dim xmlDoc As New XmlDocument()

                ' Return default settings if file doesn't exist
                If Not File.Exists(_filePath) Then
                    Return GetDefaultEmailSettings()
                End If

                xmlDoc.Load(_filePath)
                Dim emailNode = xmlDoc.SelectSingleNode("LiteTaskSettings/EmailSettings")

                If emailNode IsNot Nothing Then
                    settings("NotificationsEnabled") = GetElementValue(emailNode, "NotificationsEnabled", "False")
                    settings("SmtpServer") = GetElementValue(emailNode, "SmtpServer", "")
                    settings("SmtpPort") = GetElementValue(emailNode, "SmtpPort", "25")
                    settings("UseSSL") = GetElementValue(emailNode, "UseSSL", "True")
                    settings("EmailFrom") = GetElementValue(emailNode, "EmailFrom", "")
                    settings("EmailTo") = GetElementValue(emailNode, "EmailTo", "")
                    settings("UseCredentials") = GetElementValue(emailNode, "UseCredentials", "False")
                    settings("CredentialTarget") = GetElementValue(emailNode, "CredentialTarget", "")
                    settings("AlertLevel") = GetElementValue(emailNode, "AlertLevel", "Error")
                Else
                    Return GetDefaultEmailSettings()
                End If

                Return settings
            Catch ex As Exception
                _logger?.LogError($"Error reading email settings: {ex.Message}")
                Return GetDefaultEmailSettings()
            End Try
        End Function

        'Public Function GetLiteRunDefaults() As Dictionary(Of String, String)
        '    Try
        '        Dim defaults As New Dictionary(Of String, String)

        '        ' Check if file exists and create it with default values if not
        '        If Not File.Exists(_filePath) Then
        '            CreateDefaultConfig()
        '        End If

        '        ' Set default values
        '        defaults("Timeout") = ReadValue("LiteRunDefaults", "Timeout", "300")
        '        defaults("LogOutputPath") = ReadValue("LiteRunDefaults", "LogOutputPath", Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LiteRunLogs"))

        '        ' Ensure LiteRunDefaults section exists in XML
        '        If Not SectionExists("LiteRunDefaults") Then
        '            CreateLiteRunDefaultsSection()
        '        End If

        '        Return defaults
        '    Catch ex As Exception
        '        _logger?.LogError($"Error in GetLiteRunDefaults: {ex.Message}")
        '        Return New Dictionary(Of String, String) From {
        '    {"Timeout", "300"},
        '    {"LogOutputPath", Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LiteRunLogs")}
        '}
        '    End Try
        'End Function

        Public Function GetLogSettings() As Dictionary(Of String, String)
            Try
                Dim settings As New Dictionary(Of String, String)
                Dim xmlDoc As New XmlDocument()

                ' Return default settings if file doesn't exist
                If Not File.Exists(_filePath) Then
                    Return GetDefaultLogSettings()
                End If

                xmlDoc.Load(_filePath)
                Dim loggingNode = xmlDoc.SelectSingleNode("LiteTaskSettings/Logging")

                If loggingNode IsNot Nothing Then
                    settings("LogFolder") = GetElementValue(loggingNode, "LogFolder", Path.Combine(Application.StartupPath, "LiteTaskData", "logs"))
                    settings("LogLevel") = GetElementValue(loggingNode, "LogLevel", "Info")
                    settings("MaxLogSize") = GetElementValue(loggingNode, "MaxLogSize", "10")
                    settings("LogRetentionDays") = GetElementValue(loggingNode, "LogRetentionDays", "30")
                    settings("AlertLevel") = GetElementValue(loggingNode, "AlertLevel", "Error")
                Else
                    Return GetDefaultLogSettings()
                End If

                Return settings
            Catch ex As Exception
                _logger?.LogError($"Error reading log settings: {ex.Message}")
                Return GetDefaultLogSettings()
            End Try
        End Function

        Public Function GetToolSettings() As Dictionary(Of String, String)
            Try
                Dim settings As New Dictionary(Of String, String)

                If Not File.Exists(_filePath) Then
                    CreateDefaultConfig()
                End If

                Dim xmlDoc As New XmlDocument()
                xmlDoc.Load(_filePath)

                ' Create ToolSettings section if it doesn't exist
                Dim toolNode = xmlDoc.SelectSingleNode("LiteTaskSettings/ToolSettings")
                If toolNode Is Nothing Then
                    Dim root = xmlDoc.SelectSingleNode("LiteTaskSettings")
                    If root Is Nothing Then
                        root = xmlDoc.CreateElement("LiteTaskSettings")
                        xmlDoc.AppendChild(root)
                    End If
                    toolNode = xmlDoc.CreateElement("ToolSettings")
                    root.AppendChild(toolNode)

                    ' Add default execution tool
                    Dim execToolElem = xmlDoc.CreateElement("ExecutionTool")
                    execToolElem.InnerText = "PsExec64.exe"
                    toolNode.AppendChild(execToolElem)
                    xmlDoc.Save(_filePath)
                End If

                settings("ExecutionTool") = GetElementValue(toolNode, "ExecutionTool", "PsExec64.exe")
                Return settings

            Catch ex As Exception
                _logger?.LogError($"Error reading tool settings: {ex.Message}")
                Return New Dictionary(Of String, String) From {
            {"ExecutionTool", "PsExec64.exe"}
        }
            End Try
        End Function

        Public Function GetSqlConfiguration() As Dictionary(Of String, String)
            Try
                If Not File.Exists(_filePath) Then
                    CreateDefaultConfig()
                End If

                Dim settings As New Dictionary(Of String, String)
                Dim xmlDoc As New XmlDocument()
                xmlDoc.Load(_filePath)

                Dim sqlNode = xmlDoc.SelectSingleNode("LiteTaskSettings/SqlConfiguration")
                If sqlNode Is Nothing Then
                    sqlNode = xmlDoc.CreateElement("SqlConfiguration")
                    xmlDoc.SelectSingleNode("LiteTaskSettings").AppendChild(sqlNode)
                    sqlNode.AppendChild(CreateElement(xmlDoc, "DefaultServer", ""))
                    sqlNode.AppendChild(CreateElement(xmlDoc, "DefaultDatabase", ""))
                    sqlNode.AppendChild(CreateElement(xmlDoc, "CommandTimeout", "300"))
                    sqlNode.AppendChild(CreateElement(xmlDoc, "MaxBatchSize", "1000"))
                    xmlDoc.Save(_filePath)
                End If

                settings("DefaultServer") = GetElementValue(sqlNode, "DefaultServer", "")
                settings("DefaultDatabase") = GetElementValue(sqlNode, "DefaultDatabase", "")
                settings("CommandTimeout") = GetElementValue(sqlNode, "CommandTimeout", "300")
                settings("MaxBatchSize") = GetElementValue(sqlNode, "MaxBatchSize", "1000")

                Return settings
            Catch ex As Exception
                _logger?.LogError($"Error reading SQL configuration: {ex.Message}")
                Return New Dictionary(Of String, String) From {
            {"DefaultServer", ""},
            {"DefaultDatabase", ""},
            {"CommandTimeout", "300"},
            {"MaxBatchSize", "1000"}
        }
            End Try
        End Function

        Public Function LoadTask(taskName As String) As ScheduledTask
            If String.IsNullOrWhiteSpace(taskName) Then Return Nothing

            Try
                _xmlLock.EnterReadLock()

                Dim xmlDoc As New XmlDocument()
                xmlDoc.XmlResolver = Nothing

                If Not File.Exists(_filePath) Then
                    Return Nothing
                End If

                Using fs As New FileStream(_filePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                    xmlDoc.Load(fs)
                End Using

                Dim taskNode = xmlDoc.SelectSingleNode($"LiteTaskSettings/Tasks/Task[@Name='{SecurityElement.Escape(taskName)}']")
                If taskNode Is Nothing Then
                    Return Nothing
                End If

                Return ParseTask(taskNode)

            Catch ex As Exception
                _logger?.LogError($"Error loading task {taskName}: {ex.Message}")
                Return Nothing
            Finally
                If _xmlLock.IsReadLockHeld Then
                    _xmlLock.ExitReadLock()
                End If
            End Try
        End Function

        Private Function ParseAction(actionNode As XmlNode) As TaskAction
            Return New TaskAction With {
        .Name = GetElementValue(actionNode, "Name"),
        .Order = Integer.Parse(GetElementValue(actionNode, "Order", "1")),
        .Type = CType(Integer.Parse(GetElementValue(actionNode, "Type", "0")), TaskType),
        .Target = GetElementValue(actionNode, "Target"),
        .Parameters = GetElementValue(actionNode, "Parameters"),
        .RequiresElevation = Boolean.Parse(GetElementValue(actionNode, "RequiresElevation", "False")),
        .DependsOn = GetElementValue(actionNode, "DependsOn"),
        .WaitForCompletion = Boolean.Parse(GetElementValue(actionNode, "WaitForCompletion", "True")),
        .TimeoutMinutes = Integer.Parse(GetElementValue(actionNode, "TimeoutMinutes", "60")),
        .RetryCount = Integer.Parse(GetElementValue(actionNode, "RetryCount", "0")),
        .RetryDelayMinutes = Integer.Parse(GetElementValue(actionNode, "RetryDelayMinutes", "5")),
        .ContinueOnError = Boolean.Parse(GetElementValue(actionNode, "ContinueOnError", "False"))
    }
        End Function

        Private Function ParseTask(taskNode As XmlNode) As ScheduledTask
            Dim task As New ScheduledTask With {
            .Name = taskNode.Attributes("Name")?.Value,
            .Description = GetElementValue(taskNode, "Description"),
            .StartTime = DateTime.Parse(GetElementValue(taskNode, "StartTime", DateTime.Now.ToString("o"))),
            .Schedule = CType(Integer.Parse(GetElementValue(taskNode, "RecurrenceType", "0")), RecurrenceType),
            .Interval = TimeSpan.FromMinutes(Double.Parse(GetElementValue(taskNode, "Interval", "0"))),
            .NextRunTime = DateTime.Parse(GetElementValue(taskNode, "NextRunTime", DateTime.Now.ToString("o"))),
            .CredentialTarget = GetElementValue(taskNode, "CredentialTarget"),
            .AccountType = GetElementValue(taskNode, "AccountType", "Current User"),
            .UserSid = GetElementValue(taskNode, "UserSid"),
            .ExecutionMode = CType(Integer.Parse(GetElementValue(taskNode, "ExecutionMode", "0")), TaskExecutionMode)
        }

            ' Parse monthly properties
            Dim monthlyDayString = GetElementValue(taskNode, "MonthlyDay", "0")
            Dim monthlyDay As Integer
            If Integer.TryParse(monthlyDayString, monthlyDay) Then
                task.MonthlyDay = monthlyDay
            End If

            Dim monthlyTimeString = GetElementValue(taskNode, "MonthlyTime")
            If Not String.IsNullOrWhiteSpace(monthlyTimeString) Then
                Dim monthlyTime As TimeSpan
                If TimeSpan.TryParse(monthlyTimeString, monthlyTime) Then
                    task.MonthlyTime = monthlyTime
                End If
            End If

            ' Parse daily times
            Dim dailyTimesString = GetElementValue(taskNode, "DailyTimes")
            If Not String.IsNullOrWhiteSpace(dailyTimesString) Then
                task.DailyTimes = dailyTimesString.Split(",").
                Where(Function(t) Not String.IsNullOrWhiteSpace(t)).
                Select(Function(t) TimeSpan.Parse(t.Trim())).
                ToList()
            End If

            ' Parse next task ID
            Dim nextTaskIdString = GetElementValue(taskNode, "NextTaskId")
            If Not String.IsNullOrEmpty(nextTaskIdString) Then
                task.NextTaskId = Integer.Parse(nextTaskIdString)
            End If

            ' Parse actions
            Dim actionsNode = taskNode.SelectSingleNode("Actions")
            If actionsNode IsNot Nothing Then
                For Each actionNode As XmlNode In actionsNode.SelectNodes("Action")
                    task.Actions.Add(ParseAction(actionNode))
                Next
            End If

            Return task
        End Function

        Public Function ReadValue(section As String, key As String, defaultValue As String) As String
            Try
                Dim cacheKey = $"{section}/{key}"
                
                ' Check cache first
                If DateTime.Now < _cacheExpiry AndAlso _configCache.ContainsKey(cacheKey) Then
                    Return _configCache(cacheKey)
                End If
                
                Dim xmlDoc As New XmlDocument()
                xmlDoc.Load(_filePath)

                Dim node As XmlNode = xmlDoc.SelectSingleNode($"LiteTaskSettings/{section}/{key}")
                Dim value = If(node?.InnerText, defaultValue)
                
                ' Update cache
                _configCache(cacheKey) = value
                If _configCache.Count = 1 Then ' First item, set expiry
                    _cacheExpiry = DateTime.Now.Add(_cacheTimeout)
                End If
                
                Return value
            Catch ex As Exception
                _logger.LogError($"Error reading value for {section}/{key}: {ex.Message}")
                Return defaultValue
            End Try
        End Function

        Public Sub SaveEmailSettings(smtpServer As String, smtpPort As Integer, emailFrom As String, emailTo As String,
                               enabled As Boolean, useSSL As Boolean)
            Try
                WriteValue("EmailSettings", "NotificationsEnabled", enabled.ToString())
                WriteValue("EmailSettings", "SmtpServer", smtpServer)
                WriteValue("EmailSettings", "SmtpPort", smtpPort.ToString())
                WriteValue("EmailSettings", "UseSSL", useSSL.ToString())
                WriteValue("EmailSettings", "EmailFrom", emailFrom)
                WriteValue("EmailSettings", "EmailTo", emailTo)
            Catch ex As Exception
                _logger?.LogError($"Error saving email settings: {ex.Message}")
                Throw
            End Try
        End Sub

        Public Sub SaveLogSettings(logFolder As String, logLevel As String, maxLogSize As Integer, retentionDays As Integer)
            Try
                WriteValue("Logging", "LogFolder", logFolder)
                WriteValue("Logging", "LogLevel", logLevel)
                WriteValue("Logging", "MaxLogSize", maxLogSize.ToString())
                WriteValue("Logging", "LogRetentionDays", retentionDays.ToString())
            Catch ex As Exception
                _logger?.LogError($"Error saving log settings: {ex.Message}")
                Throw
            End Try
        End Sub

        Public Sub SaveTask(task As ScheduledTask)
            If task Is Nothing Then
                Throw New ArgumentNullException(NameOf(task))
            End If

            Try
                _xmlLock.EnterWriteLock()

                CreateBackup()

                Dim xmlDoc As New XmlDocument()
                xmlDoc.XmlResolver = Nothing ' Prevent XXE

                If Not File.Exists(_filePath) Then
                    xmlDoc.LoadXml("<LiteTaskSettings><Tasks></Tasks></LiteTaskSettings>")
                Else
                    Using fs As New FileStream(_filePath, FileMode.Open, FileAccess.Read, FileShare.None)
                        xmlDoc.Load(fs)
                    End Using
                End If

                Dim tasksNode = xmlDoc.SelectSingleNode("LiteTaskSettings/Tasks")
                If tasksNode Is Nothing Then
                    tasksNode = xmlDoc.CreateElement("Tasks")
                    xmlDoc.DocumentElement.AppendChild(tasksNode)
                End If

                ' Remove existing task if present
                Dim existingTask = tasksNode.SelectSingleNode($"Task[@Name='{SecurityElement.Escape(task.Name)}']")
                If existingTask IsNot Nothing Then
                    tasksNode.RemoveChild(existingTask)
                End If

                Dim taskElement = CreateTaskElement(xmlDoc, task)
                tasksNode.AppendChild(taskElement)

                SaveXmlSafely(xmlDoc)

            Catch ex As Exception
                _logger?.LogError($"Error saving task {task.Name}: {ex.Message}")
                Throw
            Finally
                If _xmlLock.IsWriteLockHeld Then
                    _xmlLock.ExitWriteLock()
                End If
            End Try
        End Sub

        Public Sub SaveToolSettings(executionTool As String)
            Try
                Dim xmlDoc As New XmlDocument()
                xmlDoc.Load(_filePath)

                ' Get Task nodes to preserve them
                Dim tasks = xmlDoc.SelectNodes("//Task")

                ' Create fresh structure
                xmlDoc.LoadXml("<?xml version='1.0' encoding='UTF-8'?>" &
                      "<LiteTaskSettings>" &
                      "<ToolSettings><ExecutionTool>" & executionTool & "</ExecutionTool></ToolSettings>" &
                      "<Tasks></Tasks>" &
                      "</LiteTaskSettings>")

                ' Restore tasks
                Dim tasksNode = xmlDoc.SelectSingleNode("//Tasks")
                For Each task As XmlNode In tasks
                    tasksNode.AppendChild(xmlDoc.ImportNode(task, True))
                Next

                xmlDoc.Save(_filePath)
                _logger?.LogInfo($"Tool settings saved: {executionTool}")
            Catch ex As Exception
                _logger?.LogError($"Error saving tool settings: {ex.Message}")
                Throw
            End Try
        End Sub

        'Cleans up old backup files and configuration temporary files
        Public Sub CleanupConfigFiles()
            Try
                CleanupOldBackups()
                CleanupConfigTempFiles()
            Catch ex As Exception
                _logger?.LogError($"Error during config file cleanup: {ex.Message}")
            End Try
        End Sub

        'Cleans up old backup files older than specified retention period
        Private Sub CleanupOldBackups(Optional retentionDays As Integer = 30)
            Try
                If Not System.IO.Directory.Exists(_backupPath) Then Return
                
                Dim cutoffDate = DateTime.Now.AddDays(-retentionDays)
                Dim backupFiles = System.IO.Directory.GetFiles(_backupPath, "backup_*.xml")
                
                For Each backupFile In backupFiles
                    Try
                        Dim fileInfo = New System.IO.FileInfo(backupFile)
                        If fileInfo.CreationTime < cutoffDate Then
                            System.IO.File.Delete(backupFile)
                            _logger?.LogInfo($"Deleted old backup file: {System.IO.Path.GetFileName(backupFile)}")
                        End If
                    Catch ex As Exception
                        _logger?.LogWarning($"Failed to delete backup file {backupFile}: {ex.Message}")
                    End Try
                Next
                
            Catch ex As Exception
                _logger?.LogError($"Error cleaning up old backups: {ex.Message}")
            End Try
        End Sub

        'Cleans up temporary files created during configuration operations
        Private Sub CleanupConfigTempFiles()
            Try
                If Not System.IO.Directory.Exists(_tempPath) Then Return
                
                Dim cutoffTime = DateTime.Now.AddHours(-1) ' Files older than 1 hour
                Dim tempFiles = System.IO.Directory.GetFiles(_tempPath, "xml_save_*.tmp")
                
                For Each tempFile In tempFiles
                    Try
                        Dim fileInfo = New System.IO.FileInfo(tempFile)
                        If fileInfo.LastWriteTime < cutoffTime Then
                            ' Check if file is not locked before deleting
                            If Not IsFileLocked(tempFile) Then
                                System.IO.File.Delete(tempFile)
                                _logger?.LogInfo($"Cleaned up old config temp file: {System.IO.Path.GetFileName(tempFile)}")
                            End If
                        End If
                    Catch ex As Exception
                        _logger?.LogWarning($"Failed to delete config temp file {tempFile}: {ex.Message}")
                    End Try
                Next
                
            Catch ex As Exception
                _logger?.LogError($"Error cleaning up config temp files: {ex.Message}")
            End Try
        End Sub

        'Checks if a file is currently locked by another process
        Private Function IsFileLocked(filePath As String) As Boolean
            Try
                Using fs = System.IO.File.Open(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.None)
                    Return False
                End Using
            Catch
                Return True
            End Try
        End Function

        Private Sub SaveXmlSafely(xmlDoc As XmlDocument)
            Dim tempPath = Path.Combine(_tempPath, $"xml_save_{Guid.NewGuid()}.tmp")

            Try
                Using fs As New FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None)
                    xmlDoc.Save(fs)
                End Using

                If File.Exists(_filePath) Then
                    File.Delete(_filePath)
                End If
                File.Move(tempPath, _filePath)

            Catch ex As Exception
                _logger?.LogError($"Error saving XML file safely: {ex.Message}")
                Throw
            Finally
                ' Ensure temp file is cleaned up even if an exception occurs
                If File.Exists(tempPath) Then
                    Try
                        File.Delete(tempPath)
                    Catch cleanupEx As Exception
                        _logger?.LogWarning($"Failed to cleanup temp file {tempPath}: {cleanupEx.Message}")
                    End Try
                End If
            End Try
        End Sub

        Private Function SectionExists(sectionName As String) As Boolean
            Try
                Dim xmlDoc As New XmlDocument()
                xmlDoc.Load(_filePath)
                Return xmlDoc.SelectSingleNode($"LiteTaskSettings/{sectionName}") IsNot Nothing
            Catch
                Return False
            End Try
        End Function

        Public Sub WriteValue(section As String, key As String, value As String)
            Try
                Dim xmlDoc As New XmlDocument()
                xmlDoc.Load(_filePath)

                Dim sectionNode As XmlNode = xmlDoc.SelectSingleNode($"LiteTaskSettings/{section}")
                If sectionNode Is Nothing Then
                    sectionNode = xmlDoc.CreateElement(section)
                    xmlDoc.DocumentElement.AppendChild(sectionNode)
                End If

                Dim keyNode As XmlNode = sectionNode.SelectSingleNode(key)
                If keyNode Is Nothing Then
                    keyNode = xmlDoc.CreateElement(key)
                    sectionNode.AppendChild(keyNode)
                End If

                keyNode.InnerText = value
                xmlDoc.Save(_filePath)
                
                ' Clear cache for this key and related values
                Dim cacheKey = $"{section}/{key}"
                If _configCache.ContainsKey(cacheKey) Then
                    _configCache.Remove(cacheKey)
                End If
            Catch ex As Exception
                _logger.LogError($"Error writing value for {section}/{key}: {ex.Message}")
            End Try
        End Sub

    End Class
End Namespace