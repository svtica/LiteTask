' -----------------------------------------------------------------------------
' LiteTask - part of LiteSuite. Released under The Unlicense (see LICENSE).
' File:    TaskRunReport.vb
' Author:  LiteTask contributors
' Date:    2026-11-11
' Purpose: Real timeline of one task run (task and per-action start / end /
'          duration / status) and the log lines written while it ran, rendered
'          as the body of the e-mail notification. Before this, notifications
'          only carried the 5-second batching window of the NotificationManager
'          ("Start Time" / "End Time"), so a run of several minutes looked like
'          it took a few seconds. Kept free of scheduler state so it can be
'          unit-tested.
' -----------------------------------------------------------------------------
Imports System.Collections.Concurrent
Imports System.Text

Namespace LiteTask

    ''' <summary>Status and timing of one action in a <see cref="TaskRunReport"/>.</summary>
    Public Class ActionRunInfo
        Public Property Name As String
        Public Property ActionType As String
        Public Property StartTime As DateTime?
        Public Property EndTime As DateTime?
        ''' <summary>Not run, Running, Completed, Failed, TimedOut.</summary>
        Public Property Status As String = "Not run"
        Public Property ErrorMessage As String
    End Class

    ''' <summary>
    ''' Thread-safe ring buffer of the log lines written in the execution context of
    ''' one task run (see <see cref="Logger.CurrentCapture"/>). Keeps the last
    ''' <c>maxLines</c> lines, which is where the cause of a failure usually is.
    ''' </summary>
    Public Class TaskLogCapture
        Public Const DefaultMaxLines As Integer = 500
        Public Const MaxLineLength As Integer = 1000

        Private ReadOnly _lines As New ConcurrentQueue(Of String)
        Private ReadOnly _maxLines As Integer
        Private _dropped As Integer

        Public Sub New(Optional maxLines As Integer = DefaultMaxLines)
            _maxLines = Math.Max(1, maxLines)
        End Sub

        ''' <summary>Number of older lines discarded because the buffer was full.</summary>
        Public ReadOnly Property DroppedCount As Integer
            Get
                Return Volatile.Read(_dropped)
            End Get
        End Property

        Public Sub Add(line As String)
            If line Is Nothing Then Return
            If line.Length > MaxLineLength Then
                line = line.Substring(0, MaxLineLength) & " [...]"
            End If
            _lines.Enqueue(line)
            Dim discarded As String = Nothing
            While _lines.Count > _maxLines AndAlso _lines.TryDequeue(discarded)
                Interlocked.Increment(_dropped)
            End While
        End Sub

        Public Function GetLines() As List(Of String)
            Return _lines.ToList()
        End Function
    End Class

    Public Class TaskRunReport
        Public Property TaskName As String
        Public Property StartTime As DateTime
        Public Property EndTime As DateTime?
        Public ReadOnly Property Actions As New List(Of ActionRunInfo)
        Public Property FatalError As String
        Public Property StackTrace As String
        Public ReadOnly Property Log As TaskLogCapture

        Public Sub New(taskName As String, startTime As DateTime, Optional log As TaskLogCapture = Nothing)
            Me.TaskName = taskName
            Me.StartTime = startTime
            Me.Log = If(log, New TaskLogCapture())
        End Sub

        ''' <summary>Registers every action up front so the ones never reached show as "Not run".</summary>
        Public Function AddAction(name As String, actionType As String) As ActionRunInfo
            Dim info As New ActionRunInfo With {.Name = name, .ActionType = actionType}
            Actions.Add(info)
            Return info
        End Function

        Public ReadOnly Property HasFailures As Boolean
            Get
                Return Not String.IsNullOrEmpty(FatalError) OrElse
                       Actions.Any(Function(a) a.Status = "Failed" OrElse a.Status = "TimedOut")
            End Get
        End Property

        Public ReadOnly Property Outcome As String
            Get
                If Not String.IsNullOrEmpty(FatalError) Then Return "Failed"
                If HasFailures Then Return "Completed with errors"
                If EndTime.HasValue Then Return "Succeeded"
                Return "Running"
            End Get
        End Property

        ''' <summary>"850 ms", "12.4 s", "4 min 07 s", "1 h 02 min 05 s".</summary>
        Public Shared Function FormatDuration(duration As TimeSpan) As String
            If duration < TimeSpan.Zero Then duration = TimeSpan.Zero
            If duration.TotalSeconds < 1 Then
                Return $"{CInt(Math.Floor(duration.TotalMilliseconds))} ms"
            End If
            If duration.TotalMinutes < 1 Then
                Return (Math.Floor(duration.TotalSeconds * 10) / 10).ToString("0.0", Globalization.CultureInfo.InvariantCulture) & " s"
            End If
            If duration.TotalHours < 1 Then
                Return $"{duration.Minutes} min {duration.Seconds:00} s"
            End If
            Return $"{CInt(Math.Floor(duration.TotalHours))} h {duration.Minutes:00} min {duration.Seconds:00} s"
        End Function

        ''' <summary>
        ''' Plain-text body of the notification.
        ''' </summary>
        ''' <param name="includeLog">Append the log lines captured during the run.</param>
        ''' <param name="maxLogLines">Maximum number of (most recent) log lines to append.</param>
        ''' <param name="now">End time to use when the run has not ended yet.</param>
        Public Function Format(includeLog As Boolean, maxLogLines As Integer, Optional now As DateTime? = Nothing) As String
            Dim runEnd = If(EndTime, If(now, DateTime.Now))
            Dim sb As New StringBuilder()

            sb.AppendLine($"Task:     {TaskName}")
            sb.AppendLine($"Outcome:  {Outcome}")
            sb.AppendLine($"Started:  {StartTime:yyyy-MM-dd HH:mm:ss}")
            sb.AppendLine($"Ended:    {runEnd:yyyy-MM-dd HH:mm:ss}")
            sb.AppendLine($"Duration: {FormatDuration(runEnd - StartTime)}")

            If Actions.Count > 0 Then
                sb.AppendLine()
                sb.AppendLine("Actions:")
                For i = 0 To Actions.Count - 1
                    Dim a = Actions(i)
                    Dim label = If(String.IsNullOrEmpty(a.ActionType), a.Name, $"{a.Name} ({a.ActionType})")
                    If a.StartTime.HasValue Then
                        Dim actionEnd = If(a.EndTime, runEnd)
                        sb.AppendLine($"  {i + 1}. {label} - {a.Status} - {a.StartTime.Value:HH:mm:ss} -> {actionEnd:HH:mm:ss} ({FormatDuration(actionEnd - a.StartTime.Value)})")
                    Else
                        sb.AppendLine($"  {i + 1}. {label} - {a.Status}")
                    End If
                    If Not String.IsNullOrEmpty(a.ErrorMessage) Then
                        sb.AppendLine($"     Error: {a.ErrorMessage}")
                    End If
                Next
            End If

            If Not String.IsNullOrEmpty(FatalError) Then
                sb.AppendLine()
                sb.AppendLine($"Error: {FatalError}")
                If Not String.IsNullOrEmpty(StackTrace) Then
                    sb.AppendLine("Stack trace:")
                    sb.AppendLine(StackTrace)
                End If
            End If

            If includeLog Then
                Dim lines = Log.GetLines()
                Dim limit = Math.Max(1, maxLogLines)
                Dim omitted = Log.DroppedCount + Math.Max(0, lines.Count - limit)
                If lines.Count > limit Then lines = lines.Skip(lines.Count - limit).ToList()

                sb.AppendLine()
                If lines.Count = 0 Then
                    sb.AppendLine("Task log: (no entries)")
                Else
                    sb.AppendLine($"Task log ({lines.Count} line(s){If(omitted > 0, $", {omitted} earlier line(s) omitted", "")}):")
                    For Each line In lines
                        sb.AppendLine(line)
                    Next
                End If
            End If

            Return sb.ToString()
        End Function
    End Class

End Namespace
