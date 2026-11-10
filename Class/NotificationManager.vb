Imports System.Collections.Concurrent
Imports System.Net.Mail
Imports LiteTask.LiteTask.NotificationManager

Namespace LiteTask
    Public Class NotificationManager
        Implements IDisposable

        Private _disposed As Boolean
        Private ReadOnly _logger As Logger
        Private ReadOnly _xmlManager As XMLManager
        Private _emailSettings As Dictionary(Of String, String)
        Private _smtpClient As SmtpClient
        Private _isInitialized As Boolean = False
        Private ReadOnly _messageQueue As New ConcurrentQueue(Of EmailMessage)
        Private _processingTask As Task
        Private _cancellationTokenSource As New CancellationTokenSource()
        Private _isProcessingEmails As Boolean = False
        Private ReadOnly _activeBatches As New ConcurrentDictionary(Of String, NotificationBatch)

        Public Sub New(logger As Logger, xmlManager As XMLManager)
            _logger = logger
            _xmlManager = xmlManager
            InitializeEmailSettings()
            StartEmailProcessor()
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not _disposed Then
                If disposing Then
                    Try
                        ' Cancel any pending operations
                        If _cancellationTokenSource IsNot Nothing Then
                            _cancellationTokenSource.Cancel()
                            _cancellationTokenSource.Dispose()
                            _cancellationTokenSource = Nothing
                        End If

                        ' Wait for processing task to complete
                        If _processingTask IsNot Nothing Then
                            Try
                                _processingTask.Wait(TimeSpan.FromSeconds(5))
                            Catch ex As Exception
                                _logger?.LogError($"Error waiting for processing task: {ex.Message}")
                            End Try
                            _processingTask = Nothing
                        End If

                        ' Dispose of SMTP client
                        If _smtpClient IsNot Nothing Then
                            _smtpClient.Dispose()
                            _smtpClient = Nothing
                        End If

                        ' Clear active notification batches to release held EmailMessage lists
                        _activeBatches.Clear()

                        ' Messages still queued at shutdown used to be dropped silently:
                        ' write them to the dead-letter folder instead.
                        Dim pending As EmailMessage = Nothing
                        While _messageQueue.TryDequeue(pending)
                            pending.LastError = "NotificationManager disposed (service stopping) before the message could be sent"
                            WriteDeadLetter(pending)
                        End While

                        ' Clear settings
                        If _emailSettings IsNot Nothing Then
                            _emailSettings.Clear()
                            _emailSettings = Nothing
                        End If
                    Catch ex As Exception
                        _logger?.LogError($"Error during NotificationManager disposal: {ex.Message}")
                    End Try
                End If

                _disposed = True
            End If
        End Sub

        Protected Overrides Sub Finalize()
            Dispose(disposing:=False)
            MyBase.Finalize()
        End Sub

        Private Sub InitializeEmailSettings()
            Try
                _emailSettings = _xmlManager.GetEmailSettings()
                If Boolean.Parse(_emailSettings("NotificationsEnabled")) Then
                    _smtpClient = New SmtpClient(_emailSettings("SmtpServer"), Integer.Parse(_emailSettings("SmtpPort"))) With {
                        .EnableSsl = True,
                        .DeliveryMethod = SmtpDeliveryMethod.Network,
                        .Timeout = 10000
                    }
                    _isInitialized = True
                    _logger.LogInfo("Email notification system initialized successfully")
                End If
            Catch ex As Exception
                _logger.LogError($"Failed to initialize email settings: {ex.Message}")
                _isInitialized = False
            End Try
        End Sub

        Public Sub QueueNotification(subject As String, body As String, priority As NotificationPriority)
            ThrowIfDisposed()

            Try
                If Not _isInitialized OrElse Not Boolean.Parse(_emailSettings("NotificationsEnabled")) Then
                    Return
                End If

                ' Check if message meets notification level threshold
                Dim threshold = DirectCast([Enum].Parse(GetType(Logger.LogLevel), _emailSettings("AlertLevel")), Logger.LogLevel)
                Dim messageLevel = If(priority = NotificationPriority.High,
                                Logger.LogLevel.Critical,
                                Logger.LogLevel.Error)
                If messageLevel < threshold Then
                    Return
                End If

                Dim message = New EmailMessage With {
                .Subject = subject,
                .Body = body,
                .Priority = priority,
                .Timestamp = DateTime.Now
            }

                ' Try to add to existing batch
                Dim addedToBatch = False
                For Each batch In _activeBatches.Values
                    If batch.ShouldAddToBatch(message) Then
                        batch.AddMessage(message)
                        addedToBatch = True
                        Exit For
                    End If
                Next

                ' Create new batch if needed
                If Not addedToBatch Then
                    Dim newBatch = New NotificationBatch(message)
                    _activeBatches.TryAdd(newBatch.BatchId, newBatch)

                    ' Schedule batch processing — always remove the batch from _activeBatches
                    ' to prevent accumulation, but only enqueue the email if not cancelled.
                    Task.Delay(5000, _cancellationTokenSource.Token).ContinueWith(Sub(t)
                                                      Dim batch As NotificationBatch = Nothing
                                                      If _activeBatches.TryRemove(newBatch.BatchId, batch) Then
                                                          If Not t.IsCanceled Then
                                                              Dim batchedMessage = New EmailMessage With {
                            .Subject = If(batch.Messages.Count > 1,
                                         $"{batch.Subject} - {batch.Messages.Count} messages",
                                         batch.Subject),
                            .Body = batch.GetCombinedBody(),
                            .Priority = batch.HighestPriority,
                            .Timestamp = batch.StartTime
                        }
                                                              _messageQueue.Enqueue(batchedMessage)
                                                          End If
                                                          batch.Messages.Clear()
                                                      End If
                                                  End Sub)
                End If

                _logger.LogInfo($"Email notification queued: {subject}")
            Catch ex As Exception
                _logger.LogError($"Error queueing notification: {ex.Message}")
            End Try
        End Sub

        Private Sub ThrowIfDisposed()
            If _disposed Then
                Throw New ObjectDisposedException(GetType(NotificationManager).FullName)
            End If
        End Sub

        Private Sub StartEmailProcessor()
            _processingTask = Task.Run(Async Function()
                                           While Not _cancellationTokenSource.Token.IsCancellationRequested
                                               Try
                                                   Dim message As EmailMessage = Nothing
                                                   If Not _messageQueue.TryDequeue(message) Then
                                                       _isProcessingEmails = False
                                                       Await Task.Delay(1000, _cancellationTokenSource.Token)
                                                       Continue While
                                                   End If

                                                   ' A message waiting for its retry window goes back to the end
                                                   ' of the queue so it never holds up the messages behind it.
                                                   If message.NotBefore > DateTime.Now Then
                                                       _messageQueue.Enqueue(message)
                                                       Await Task.Delay(1000, _cancellationTokenSource.Token)
                                                       Continue While
                                                   End If

                                                   _isProcessingEmails = True
                                                   Dim sendError As Exception = Nothing
                                                   Try
                                                       Await SendAsync(message)
                                                       _logger.LogInfo($"Email sent successfully: {message.Subject}{If(message.RetryCount > 0, $" (after {message.RetryCount} retry/retries)", "")}")
                                                   Catch ex As Exception
                                                       ' Handled after the Try: VB does not allow Await in a Catch block.
                                                       sendError = ex
                                                   End Try

                                                   If sendError IsNot Nothing Then
                                                       HandleSendFailure(message, sendError)
                                                   End If
                                                   _isProcessingEmails = False
                                               Catch ex As Exception When TypeOf ex Is OperationCanceledException
                                                   ' Normal cancellation
                                                   Exit While
                                               Catch ex As Exception
                                                   _logger.LogError($"Error in email processor: {ex.Message}")
                                                   _isProcessingEmails = False
                                               End Try
                                           End While
                                       End Function, _cancellationTokenSource.Token)
        End Sub

        Private Async Function SendAsync(message As EmailMessage) As Task
            Using mail As New MailMessage()
                mail.From = New MailAddress(_emailSettings("EmailFrom"))
                For Each recipient In _emailSettings("EmailTo").Split(";"c)
                    If Not String.IsNullOrWhiteSpace(recipient) Then
                        mail.To.Add(recipient.Trim())
                    End If
                Next
                mail.Subject = message.Subject
                mail.Body = $"{message.Body}{Environment.NewLine}{Environment.NewLine}Timestamp: {message.Timestamp}"
                mail.Priority = If(message.Priority = NotificationPriority.High,
                                   MailPriority.High, MailPriority.Normal)

                ' One SmtpClient per send: the instance created in InitializeEmailSettings
                ' is deliberately not reused (SmtpClient is not safe for concurrent or
                ' repeated async sends and would pin a stale connection).
                Using smtp As New SmtpClient(_emailSettings("SmtpServer"), Integer.Parse(_emailSettings("SmtpPort"))) With {
                    .EnableSsl = Boolean.Parse(_emailSettings("UseSSL")),
                    .DeliveryMethod = SmtpDeliveryMethod.Network,
                    .Timeout = 30000
                }
                    Await smtp.SendMailAsync(mail)
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Schedules a retry (1 / 5 / 15 min) or, once retries are exhausted,
        ''' writes the message to the dead-letter folder so it is never lost
        ''' silently. The logged cause is the innermost exception: SmtpException
        ''' only says "Failure sending mail.", the DNS/socket/TLS reason is inside.
        ''' </summary>
        Private Sub HandleSendFailure(message As EmailMessage, ex As Exception)
            Dim cause = ex.GetBaseException().Message
            message.LastError = If(String.Equals(cause, ex.Message, StringComparison.Ordinal), ex.Message, $"{ex.Message} ({cause})")

            If NotificationRetryPolicy.ShouldRetry(message.RetryCount) Then
                message.RetryCount += 1
                Dim delay = NotificationRetryPolicy.GetRetryDelay(message.RetryCount)
                message.NotBefore = DateTime.Now.Add(delay)
                _messageQueue.Enqueue(message)
                _logger.LogError($"Failed to send email '{message.Subject}': {message.LastError}. Retry {message.RetryCount}/{NotificationRetryPolicy.MaxRetries} in {delay.TotalMinutes:F0} min")
            Else
                _logger.LogError($"Failed to send email '{message.Subject}': {message.LastError}. Giving up after {message.RetryCount + 1} attempts; message written to the dead-letter folder")
                WriteDeadLetter(message)
            End If
        End Sub

        ''' <summary>
        ''' LiteTaskData\notifications\failed\yyyyMMdd_HHmmss_&lt;guid&gt;.txt: subject,
        ''' body, attempts and last error of a notification that could not be sent.
        ''' </summary>
        Private Sub WriteDeadLetter(message As EmailMessage)
            Try
                Dim folder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LiteTaskData", "notifications", "failed")
                Directory.CreateDirectory(folder)
                Dim filePath = Path.Combine(folder, NotificationRetryPolicy.BuildDeadLetterFileName(DateTime.Now, Guid.NewGuid()))
                File.WriteAllText(filePath, NotificationRetryPolicy.FormatDeadLetter(message.Subject, message.Body, message.Timestamp, message.RetryCount + 1, message.LastError))
                _logger?.LogWarning($"Undelivered notification saved to {filePath}")
            Catch ioEx As Exception
                _logger?.LogError($"Could not write dead-letter file for '{message.Subject}': {ioEx.Message}")
            End Try
        End Sub

        Public Enum NotificationPriority
            Normal
            High
        End Enum

    End Class

    Public Class EmailMessage
        Public Property Subject As String
        Public Property Body As String
        Public Property Priority As NotificationPriority
        Public Property Timestamp As DateTime
        Public Property RetryCount As Integer = 0
        ''' <summary>Earliest time the next send attempt may run (retry backoff).</summary>
        Public Property NotBefore As DateTime = DateTime.MinValue
        ''' <summary>Reason of the last failed attempt, kept for the dead-letter file.</summary>
        Public Property LastError As String
    End Class

    Public Class NotificationBatch
        Public Property Messages As New List(Of EmailMessage)
        Public Property StartTime As DateTime
        Public Property Subject As String
        Public Property HighestPriority As NotificationPriority = NotificationPriority.Normal
        Public Property BatchId As String

        Public Sub New(initialMessage As EmailMessage)
            StartTime = DateTime.Now
            Messages = New List(Of EmailMessage) From {initialMessage}
            Subject = initialMessage.Subject
            HighestPriority = initialMessage.Priority
            BatchId = Guid.NewGuid().ToString()
        End Sub

        Public Function ShouldAddToBatch(message As EmailMessage) As Boolean
            ' Messages are related if they have the same subject (excluding timestamps)
            ' and are within 5 seconds of each other
            Return message.Subject.Split(" - ")(0) = Subject.Split(" - ")(0) AndAlso
                   DateTime.Now.Subtract(StartTime).TotalSeconds <= 5
        End Function

        Public Sub AddMessage(message As EmailMessage)
            Messages.Add(message)
            If message.Priority > HighestPriority Then
                HighestPriority = message.Priority
            End If
        End Sub

        Public Function GetCombinedBody() As String
            Dim body As New StringBuilder()

            ' Add execution summary
            body.AppendLine($"Execution Summary:")
            body.AppendLine($"Start Time: {StartTime:yyyy-MM-dd HH:mm:ss}")
            body.AppendLine($"End Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
            body.AppendLine($"Total Messages: {Messages.Count}")
            body.AppendLine()

            ' Group messages by type
            Dim outputs = Messages.Where(Function(m) Not m.Body.StartsWith("Error:")).ToList()
            Dim errors = Messages.Where(Function(m) m.Body.StartsWith("Error:")).ToList()

            If outputs.Count > 0 Then
                body.AppendLine("Execution Output:")
                For Each msg In outputs
                    body.AppendLine(msg.Body)
                Next
                body.AppendLine()
            End If

            If errors.Count > 0 Then
                body.AppendLine("Errors:")
                For Each msg In errors
                    body.AppendLine(msg.Body)
                Next
                body.AppendLine()
            End If

            Return body.ToString()
        End Function
    End Class

End Namespace
