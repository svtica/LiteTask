Namespace LiteTask
    Public Class ScheduledTask
        Public Enum TaskType
            PowerShell
            Batch
            SQL
            RemoteExecution
            Executable
        End Enum

        Public Enum RecurrenceType
            OneTime
            Interval
            Daily
        End Enum

        Public Enum TaskExecutionMode
            Auto
            Local
            Remote
            CredSSP
            SecureChannel
            PSRemoting
            OSQL
        End Enum
        ' Basic properties
        Public Property Name As String
        Public Property Description As String

        ' Scheduling properties
        Public Property StartTime As DateTime
        Public Property Schedule As RecurrenceType
        Public Property Interval As TimeSpan
        Public Property DailyTimes As List(Of TimeSpan)
        Public Property NextRunTime As DateTime

        ' Credential properties
        Public Property CredentialTarget As String
        Public Property AccountType As String
        Public Property UserSid As String

        ' Execution properties
        Public Property Actions As List(Of TaskAction)
        Public Property NextTaskId As Integer?
        Public Property ExecutionMode As TaskExecutionMode
        Public Property ServiceAccount As CredentialManager.ServiceAccountType?
        Public Property FilePath As String
        Public Property Arguments As String
        Public Property RequiresElevation As Boolean
        Public Property Parameters As Hashtable

        Public Sub New()
            DailyTimes = New List(Of TimeSpan)
            Actions = New List(Of TaskAction)
            Parameters = New Hashtable()
        End Sub

        Public Sub New(name As String, description As String, type As TaskType, filePath As String, arguments As String,
                   startTime As DateTime, recurrenceType As RecurrenceType, interval As TimeSpan,
                   dailyTimes As List(Of TimeSpan), credentialTarget As String, requiresElevation As Boolean,
                   isRemoteExecution As Boolean, remoteServerUrl As String, executionMode As TaskExecutionMode,
                   nextTaskId As Integer?)

            Me.New()  ' Call default constructor to initialize collections

            Me.Name = name
            Me.Description = description
            Me.StartTime = startTime
            Me.Schedule = recurrenceType
            Me.Interval = interval
            Me.DailyTimes = dailyTimes
            Me.CredentialTarget = credentialTarget
            Me.NextTaskId = nextTaskId
            Me.ExecutionMode = executionMode

            ' Create initial action from the legacy parameters
            Dim initialAction = New TaskAction With {
                .Order = 1,
                .Type = type,
                .Target = filePath,
                .Parameters = arguments,
                .RequiresElevation = requiresElevation
            }
            Me.Actions.Add(initialAction)

            ' Set initial NextRunTime
            Me.NextRunTime = CalculateNextRunTime()
        End Sub

        Public Function CalculateNextRunTime() As DateTime
            Dim now = DateTime.Now
            Select Case Schedule
                Case RecurrenceType.OneTime
                    If StartTime > now Then
                        Return StartTime
                    Else
                        Return DateTime.MaxValue ' Indicates that the one-time task has already run
                    End If

                Case RecurrenceType.Interval
                    If NextRunTime = DateTime.MinValue OrElse NextRunTime <= now Then
                        ' Calculate the next run time based on the current time and interval
                        Dim timeSinceStart = now - StartTime
                        Dim intervalsElapsed = Math.Ceiling(timeSinceStart.TotalMinutes / Interval.TotalMinutes)
                        Return StartTime.AddMinutes(intervalsElapsed * Interval.TotalMinutes)
                    Else
                        Return NextRunTime
                    End If

                Case RecurrenceType.Daily
                    If DailyTimes.Count = 0 Then
                        Return now.Date.AddDays(1) ' Default to next day if no daily times are set
                    Else
                        Dim todayTimes = DailyTimes.Select(Function(t) now.Date.Add(t))
                        Dim nextTime = todayTimes.FirstOrDefault(Function(t) t > now)
                        If nextTime = DateTime.MinValue Then
                            ' If all times for today have passed, get the first time for tomorrow
                            nextTime = now.Date.AddDays(1).Add(DailyTimes.Min())
                        End If
                        Return nextTime
                    End If

                Case Else
                    Throw New NotImplementedException("Unsupported recurrence type")
            End Select
        End Function

        Public Function Clone() As ScheduledTask
            Return New ScheduledTask With {
        .Name = Me.Name,
        .Description = Me.Description,
        .StartTime = Me.StartTime,
        .Schedule = Me.Schedule,
        .Interval = Me.Interval,
        .DailyTimes = New List(Of TimeSpan)(Me.DailyTimes),
        .CredentialTarget = Me.CredentialTarget,
        .AccountType = Me.AccountType,
        .RequiresElevation = Me.RequiresElevation,
        .Actions = Me.Actions.Select(Function(a) a.Clone()).ToList()
    }
        End Function

        Public Sub UpdateNextRunTime()
            Dim now = DateTime.Now
            Select Case Schedule
                Case RecurrenceType.OneTime
                    If NextRunTime <= now Then
                        NextRunTime = DateTime.MaxValue ' Indicates that the task has run
                    End If

                Case RecurrenceType.Interval
                    While NextRunTime <= now
                        NextRunTime = NextRunTime.Add(Interval)
                    End While

                Case RecurrenceType.Daily
                    NextRunTime = CalculateNextRunTime()
            End Select
        End Sub

    End Class
End Namespace