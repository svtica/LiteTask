Namespace LiteTask
    Public Class SettingsValidator

        ' Wired into OptionsForm.ValidateSettings (Forms/OptionsForm.vb), called from
        ' SaveSettings, OnFormClosing, and TestEmailButton_Click.

        Private ReadOnly _logger As Logger

        Public Sub New(logger As Logger)
            _logger = logger
        End Sub

        Public Function ValidateEmailSettings(settings As Dictionary(Of String, String)) As List(Of String)
            Dim errors As New List(Of String)

            Try
                If Boolean.Parse(settings("NotificationsEnabled")) Then
                    ' SMTP Server
                    If String.IsNullOrWhiteSpace(settings("SmtpServer")) Then
                        errors.Add("SMTP Server is required when notifications are enabled")
                    End If

                    ' SMTP Port
                    Dim port As Integer
                    If Not Integer.TryParse(settings("SmtpPort"), port) OrElse port < 1 OrElse port > 65535 Then
                        errors.Add("SMTP Port must be between 1 and 65535")
                    End If

                    ' Email Addresses
                    If String.IsNullOrWhiteSpace(settings("EmailFrom")) Then
                        errors.Add("From email address is required")
                    ElseIf Not IsValidEmail(settings("EmailFrom")) Then
                        errors.Add("Invalid From email address format")
                    End If

                    If String.IsNullOrWhiteSpace(settings("EmailTo")) Then
                        errors.Add("To email address is required")
                    Else
                        For Each email In settings("EmailTo").Split(";"c)
                            If Not IsValidEmail(email.Trim()) Then
                                errors.Add($"Invalid To email address format: {email}")
                            End If
                        Next
                    End If
                End If
            Catch ex As Exception
                _logger.LogError("Error validating email settings", ex)
                errors.Add($"Error validating email settings: {ex.Message}")
            End Try

            Return errors
        End Function

        Public Function ValidateLogSettings(settings As Dictionary(Of String, String)) As List(Of String)
            Dim errors As New List(Of String)

            Try
                ' Log Folder
                If String.IsNullOrWhiteSpace(settings("LogFolder")) Then
                    errors.Add("Log folder path is required")
                Else
                    Try
                        ' Check if path is valid and accessible
                        Dim testPath = Path.GetFullPath(settings("LogFolder"))
                        If Not Directory.Exists(testPath) Then
                            Directory.CreateDirectory(testPath)
                        End If
                        ' Test write access
                        Dim testFile = Path.Combine(testPath, "test.log")
                        File.WriteAllText(testFile, "Test")
                        File.Delete(testFile)
                    Catch ex As Exception
                        errors.Add($"Log folder is not accessible: {ex.Message}")
                    End Try
                End If

                ' Log Level
                If Not [Enum].TryParse(Of Logger.LogLevel)(settings("LogLevel"), True, Nothing) Then
                    errors.Add("Invalid log level specified")
                End If

                ' Max Log Size
                Dim maxSize As Integer
                If Not Integer.TryParse(settings("MaxLogSize"), maxSize) OrElse maxSize < 1 Then
                    errors.Add("Max log size must be greater than 0 MB")
                End If

                ' Retention Days
                Dim retentionDays As Integer
                If Not Integer.TryParse(settings("LogRetentionDays"), retentionDays) OrElse retentionDays < 1 Then
                    errors.Add("Log retention days must be greater than 0")
                End If

            Catch ex As Exception
                _logger.LogError("Error validating log settings", ex)
                errors.Add($"Error validating log settings: {ex.Message}")
            End Try

            Return errors
        End Function

        Public Function ValidateTaskAction(action As TaskAction) As List(Of String)
            Return ValidateTaskAction(action, Nothing)
        End Function

        ' Validates a single TaskAction. When siblingActions is provided, DependsOn
        ' is checked against the Names in that collection (same-task semantics, see
        ' TaskDependencyManager). PowerShell modules are intentionally not preflighted.
        Public Function ValidateTaskAction(action As TaskAction, siblingActions As IEnumerable(Of TaskAction)) As List(Of String)
            Dim errors As New List(Of String)

            Try
                If action Is Nothing Then
                    errors.Add("Action is null")
                    Return errors
                End If

                If String.IsNullOrWhiteSpace(action.Name) Then
                    errors.Add("Action name is required")
                End If

                If String.IsNullOrWhiteSpace(action.Target) Then
                    errors.Add("Action target is required")
                Else
                    Select Case action.Type
                        Case ScheduledTask.TaskType.PowerShell,
                             ScheduledTask.TaskType.Batch,
                             ScheduledTask.TaskType.Executable
                            If Not File.Exists(action.Target) Then
                                errors.Add($"Target file does not exist: {action.Target}")
                            End If
                        Case ScheduledTask.TaskType.SQL
                            ' Target holds the SQL string or path; require non-empty (already covered above)
                        Case ScheduledTask.TaskType.RemoteExecution
                            ' Target is a remote endpoint or script path; non-empty is sufficient here
                    End Select
                End If

                If action.TimeoutMinutes < 1 Then
                    errors.Add("Timeout must be at least 1 minute")
                End If

                If action.RetryCount < 0 Then
                    errors.Add("Retry count cannot be negative")
                End If

                If action.RetryDelayMinutes < 0 Then
                    errors.Add("Retry delay cannot be negative")
                End If

                If Not String.IsNullOrWhiteSpace(action.DependsOn) AndAlso siblingActions IsNot Nothing Then
                    Dim selfName = If(action.Name, String.Empty)
                    Dim known = siblingActions _
                        .Where(Function(a) a IsNot Nothing AndAlso Not Object.ReferenceEquals(a, action)) _
                        .Select(Function(a) a.Name) _
                        .Where(Function(n) Not String.IsNullOrWhiteSpace(n))

                    If String.Equals(action.DependsOn, selfName, StringComparison.OrdinalIgnoreCase) Then
                        errors.Add($"Action '{selfName}' cannot depend on itself")
                    ElseIf Not known.Any(Function(n) String.Equals(n, action.DependsOn, StringComparison.OrdinalIgnoreCase)) Then
                        errors.Add($"DependsOn references unknown action: {action.DependsOn}")
                    End If
                End If

            Catch ex As Exception
                _logger?.LogError("Error validating task action", ex)
                errors.Add($"Error validating task action: {ex.Message}")
            End Try

            Return errors
        End Function

        Private Function IsValidEmail(email As String) As Boolean
            Try
                Dim addr = New System.Net.Mail.MailAddress(email)
                Return addr.Address = email
            Catch
                Return False
            End Try
        End Function

    End Class
End Namespace