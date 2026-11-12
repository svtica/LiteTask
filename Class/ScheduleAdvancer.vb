' -----------------------------------------------------------------------------
' LiteTask - part of LiteSuite. Released under The Unlicense (see LICENSE).
' File:    ScheduleAdvancer.vb
' Author:  LiteTask contributors
' Date:    2026-11-10
' Purpose: Advance the schedule of a task that just finished, including the
'          reloaded copy when a hot reload (CMD_RELOAD_TASKS) replaced the
'          in-memory definition during the run. Kept out of CustomScheduler so
'          the two-object case can be unit-tested (GLPI #4302).
'
'          Before: CustomScheduler.ReloadTasks replaced every ScheduledTask in
'          _tasks, including the one executing. At the end of that run only the
'          orphaned copy was advanced while SaveTasks persisted the reloaded
'          copy with its pre-run NextRunTime, so the task was dispatched again
'          on the next tick.
' -----------------------------------------------------------------------------
Namespace LiteTask
    Public Module ScheduleAdvancer

        ''' <summary>
        ''' Calls UpdateNextRunTime on the executed task and, when
        ''' <paramref name="liveTask"/> is a different object (hot reload during
        ''' the run), on that one too. Returns True when the live copy was advanced.
        ''' </summary>
        Public Function Advance(executedTask As ScheduledTask, liveTask As ScheduledTask) As Boolean
            If executedTask Is Nothing Then Throw New ArgumentNullException(NameOf(executedTask))
            executedTask.UpdateNextRunTime()
            If liveTask IsNot Nothing AndAlso liveTask IsNot executedTask Then
                liveTask.UpdateNextRunTime()
                Return True
            End If
            Return False
        End Function

    End Module
End Namespace
