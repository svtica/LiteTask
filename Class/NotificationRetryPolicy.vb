' -----------------------------------------------------------------------------
' Copyright (c) svtica. All rights reserved.
' File:    NotificationRetryPolicy.vb
' Author:  LiteTask contributors
' Date:    2026-11-09
' Purpose: Retry schedule and dead-letter format for e-mail notifications.
'          Kept out of NotificationManager so it can be unit-tested without
'          an SMTP server.
'
'          GLPI #4302: on 2026-10-27 12:01 two alerts failed 4 times each
'          within one second (immediate retries, v1.2.1.9) and were dropped
'          without any trace; the SMTP relay was reachable again 30 minutes
'          later. Retries are now spread over ~21 minutes and a message that
'          still cannot be sent is written to LiteTaskData\notifications\failed.
' -----------------------------------------------------------------------------
Namespace LiteTask
    Public Module NotificationRetryPolicy

        ''' <summary>Number of retries after the first attempt.</summary>
        Public Const MaxRetries As Integer = 3

        Private ReadOnly RetryDelays As TimeSpan() = {
            TimeSpan.FromMinutes(1),
            TimeSpan.FromMinutes(5),
            TimeSpan.FromMinutes(15)
        }

        ''' <summary>True while another retry is allowed for a message that has already been retried <paramref name="retryCount"/> times.</summary>
        Public Function ShouldRetry(retryCount As Integer) As Boolean
            Return retryCount < MaxRetries
        End Function

        ''' <summary>Delay before retry number <paramref name="retryNumber"/> (1-based). Later retries reuse the last delay.</summary>
        Public Function GetRetryDelay(retryNumber As Integer) As TimeSpan
            If retryNumber < 1 Then Return RetryDelays(0)
            Return RetryDelays(Math.Min(retryNumber, RetryDelays.Length) - 1)
        End Function

        ''' <summary>Total time a message can stay in the retry cycle before it is dead-lettered.</summary>
        Public Function TotalRetryWindow() As TimeSpan
            Dim total = TimeSpan.Zero
            For i = 1 To MaxRetries
                total = total.Add(GetRetryDelay(i))
            Next
            Return total
        End Function

        Public Function BuildDeadLetterFileName(timestamp As DateTime, id As Guid) As String
            Return $"{timestamp:yyyyMMdd_HHmmss}_{id:N}.txt"
        End Function

        ''' <summary>Plain-text content of a dead-letter file (readable by a human, re-sendable by hand).</summary>
        Public Function FormatDeadLetter(subject As String, body As String, timestamp As DateTime, attempts As Integer, lastError As String) As String
            Dim sb As New StringBuilder()
            sb.AppendLine("LiteTask undelivered e-mail notification")
            sb.AppendLine($"Queued:    {timestamp:yyyy-MM-dd HH:mm:ss}")
            sb.AppendLine($"Attempts:  {attempts}")
            sb.AppendLine($"LastError: {If(lastError, String.Empty)}")
            sb.AppendLine($"Subject:   {If(subject, String.Empty)}")
            sb.AppendLine()
            sb.AppendLine(If(body, String.Empty))
            Return sb.ToString()
        End Function

    End Module
End Namespace
