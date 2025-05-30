Namespace LiteTask
    Public Class SettingsValidator

        'TODO Reimplement in Options

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