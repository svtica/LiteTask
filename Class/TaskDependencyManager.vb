Imports LiteTask.LiteTask.ScheduledTask

Namespace LiteTask

    Public Class TaskDependencyManager
        Private ReadOnly _logger As Logger
        Private ReadOnly _xmlManager As XMLManager
        Private ReadOnly _customScheduler As CustomScheduler

        Public Sub New(logger As Logger, xmlManager As XMLManager, customScheduler As CustomScheduler)
            _logger = logger
            _xmlManager = xmlManager
            _customScheduler = customScheduler
        End Sub

        Private Function HasCircularDependency(taskName As String, visitedTasks As HashSet(Of String)) As Boolean
            If visitedTasks.Contains(taskName) Then
                Return True
            End If

            visitedTasks.Add(taskName)
            Dim task = _customScheduler.GetTask(taskName)

            If task IsNot Nothing Then
                For Each action In task.Actions
                    If action.DependsOn IsNot Nothing Then
                        If HasCircularDependency(action.DependsOn, New HashSet(Of String)(visitedTasks)) Then
                            Return True
                        End If
                    End If
                Next
            End If

            visitedTasks.Remove(taskName)
            Return False
        End Function

        Private Function TopologicalSort(
            action As TaskAction,
            task As ScheduledTask,
            visited As HashSet(Of String),
            processing As HashSet(Of String),
            sortedActions As List(Of TaskAction)) As Boolean

            If processing.Contains(action.Name) Then
                Return False ' Circular dependency
            End If

            If visited.Contains(action.Name) Then
                Return True ' Already processed
            End If

            processing.Add(action.Name)

            ' Process dependencies
            If action.DependsOn IsNot Nothing Then
                Dim dependentAction = task.Actions.FirstOrDefault(Function(a) a.Name = action.DependsOn)
                If dependentAction IsNot Nothing Then
                    If Not TopologicalSort(dependentAction, task, visited, processing, sortedActions) Then
                        Return False
                    End If
                End If
            End If

            processing.Remove(action.Name)
            visited.Add(action.Name)
            sortedActions.Add(action)
            Return True
        End Function

    End Class

    Public Class TaskAction

        Public Property Order As Integer
        Public Property Name As String
        Public Property Type As TaskType
        Public Property Target As String
        Public Property Parameters As String
        Public Property RequiresElevation As Boolean
        Public Property DependsOn As String
        Public Property WaitForCompletion As Boolean = True
        Public Property TimeoutMinutes As Integer = 60
        Public Property RetryCount As Integer = 0
        Public Property RetryDelayMinutes As Integer = 5
        Public Property ContinueOnError As Boolean = False
        Public Property Status As TaskActionStatus = TaskActionStatus.Pending
        Public Property LastRunTime As DateTime?
        Public Property NextRetryTime As DateTime?

        Public Function Clone() As TaskAction
            Return DirectCast(Me.MemberwiseClone(), TaskAction)
        End Function
    End Class

    Public Enum TaskActionStatus
        Pending
        Running
        Completed
        Failed
        Retrying
        TimedOut
        Skipped
    End Enum

End Namespace