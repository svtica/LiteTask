' -----------------------------------------------------------------------------
' LiteTask - part of LiteSuite. Released under The Unlicense (see LICENSE).
' File:    ScheduleAdvancerTests.vb
' Author:  LiteTask contributors
' Date:    2026-11-10
' Purpose: Unit tests for LiteTask.ScheduleAdvancer: a hot reload
'          during a run must not leave the reloaded task object with its
'          pre-run NextRunTime (which re-dispatched the task on the next tick).
' -----------------------------------------------------------------------------
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports LiteTask.LiteTask

<TestClass>
Public Class ScheduleAdvancerTests

    ' Interval task whose NextRunTime is in the past: the "due" state a task has
    ' while it executes, before the scheduler advances it.
    ' Two copies of the same task (executed and reloaded) must be built from the same
    ' reference time: two DateTime.Now calls differ by a few ticks, and so would their
    ' next occurrences.
    Private Shared Function DueIntervalTask(name As String) As ScheduledTask
        Return DueIntervalTask(name, DateTime.Now)
    End Function

    Private Shared Function DueIntervalTask(name As String, reference As DateTime) As ScheduledTask
        Dim t As New ScheduledTask()
        t.Name = name
        t.Schedule = ScheduledTask.RecurrenceType.Interval
        t.Interval = TimeSpan.FromMinutes(30)
        t.StartTime = reference.AddHours(-5)
        t.NextRunTime = reference.AddMinutes(-2)
        Return t
    End Function

    <TestMethod>
    Public Sub Advance_SameObject_AdvancesOnceAndReportsNoLiveCopy()
        Dim task = DueIntervalTask("SqlExport")
        Dim advancedLive = ScheduleAdvancer.Advance(task, task)
        Assert.IsFalse(advancedLive)
        Assert.IsTrue(task.NextRunTime > DateTime.Now, "executed task must be advanced past now")
    End Sub

    <TestMethod>
    Public Sub Advance_NoLiveCopy_AdvancesExecutedOnly()
        Dim task = DueIntervalTask("SqlExport")
        Assert.IsFalse(ScheduleAdvancer.Advance(task, Nothing))
        Assert.IsTrue(task.NextRunTime > DateTime.Now)
    End Sub

    <TestMethod>
    Public Sub Advance_ReloadedCopy_IsAdvancedToo()
        ' Simulates ReloadTasks replacing _tasks("SqlExport") while the run is in progress:
        ' the executing code holds 'executed', the scheduler dictionary holds 'reloaded'.
        Dim reference As DateTime = DateTime.Now
        Dim executed = DueIntervalTask("SqlExport", reference)
        Dim reloaded = DueIntervalTask("SqlExport", reference)
        Assert.IsTrue(reloaded.NextRunTime <= DateTime.Now, "precondition: reloaded copy is still due")

        Dim advancedLive = ScheduleAdvancer.Advance(executed, reloaded)

        Assert.IsTrue(advancedLive)
        Assert.IsTrue(executed.NextRunTime > DateTime.Now)
        Assert.IsTrue(reloaded.NextRunTime > DateTime.Now, "reloaded copy must not stay due, or the next tick re-dispatches the task")
        Assert.AreEqual(executed.NextRunTime, reloaded.NextRunTime, "same schedule => same next occurrence")
    End Sub

    <TestMethod>
    Public Sub Advance_NullExecutedTask_Throws()
        Assert.ThrowsException(Of ArgumentNullException)(Sub() ScheduleAdvancer.Advance(Nothing, Nothing))
    End Sub

End Class
