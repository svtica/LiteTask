' -----------------------------------------------------------------------------
' LiteTask - part of LiteSuite. Released under The Unlicense (see LICENSE).
' File:    TaskRunReportTests.vb
' Author:  LiteTask contributors
' Date:    2026-11-11
' Purpose: Unit tests for LiteTask.TaskRunReport / TaskLogCapture: the e-mail
'          notification must carry the real start, end and duration of the run
'          (not the 5-second batching window) and, when enabled, the log lines
'          captured during that run.
' -----------------------------------------------------------------------------
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports LiteTask.LiteTask

<TestClass>
Public Class TaskRunReportTests

    Private Shared ReadOnly T0 As New DateTime(2026, 11, 11, 2, 0, 0)

    <TestMethod>
    Public Sub Format_ShowsRealStartEndAndDuration()
        Dim report As New TaskRunReport("Export", T0)
        Dim a1 = report.AddAction("Extract", "SQL")
        a1.StartTime = T0
        a1.EndTime = T0.AddMinutes(5).AddSeconds(10)
        a1.Status = "Completed"
        report.EndTime = T0.AddMinutes(7).AddSeconds(43)

        Dim body = report.Format(False, 200)

        StringAssert.Contains(body, "Started:  2026-11-11 02:00:00")
        StringAssert.Contains(body, "Ended:    2026-11-11 02:07:43")
        StringAssert.Contains(body, "Duration: 7 min 43 s")
        StringAssert.Contains(body, "1. Extract (SQL) - Completed - 02:00:00 -> 02:05:10 (5 min 10 s)")
        StringAssert.Contains(body, "Outcome:  Succeeded")
    End Sub

    <TestMethod>
    Public Sub Format_FailedRun_ListsErrorAndActionsNotRun()
        Dim report As New TaskRunReport("Export", T0)
        Dim a1 = report.AddAction("Extract", "SQL")
        report.AddAction("Upload", "PowerShell")
        a1.StartTime = T0
        a1.EndTime = T0.AddMinutes(3)
        a1.Status = "Failed"
        a1.ErrorMessage = "Process exited with code 1"
        report.EndTime = T0.AddMinutes(3)
        report.FatalError = "Action failed: Extract"

        Dim body = report.Format(False, 200)

        StringAssert.Contains(body, "Outcome:  Failed")
        StringAssert.Contains(body, "Error: Process exited with code 1")
        StringAssert.Contains(body, "2. Upload (PowerShell) - Not run")
        StringAssert.Contains(body, "Error: Action failed: Extract")
    End Sub

    <TestMethod>
    Public Sub Outcome_ContinueOnErrorFailure_IsCompletedWithErrors()
        Dim report As New TaskRunReport("Export", T0)
        report.AddAction("Extract", "SQL").Status = "Failed"
        report.EndTime = T0.AddMinutes(1)
        Assert.IsTrue(report.HasFailures)
        Assert.AreEqual("Completed with errors", report.Outcome)
    End Sub

    <TestMethod>
    Public Sub Format_Log_OnlyWhenRequested_AndLimitedToLastLines()
        Dim capture As New TaskLogCapture()
        For i = 1 To 10
            capture.Add($"line {i}")
        Next
        Dim report As New TaskRunReport("Export", T0, capture) With {.EndTime = T0.AddSeconds(30)}

        Assert.IsFalse(report.Format(False, 200).Contains("Task log"))

        Dim body = report.Format(True, 3)
        StringAssert.Contains(body, "Task log (3 line(s), 7 earlier line(s) omitted):")
        StringAssert.Contains(body, "line 10")
        Assert.IsFalse(body.Contains("line 7" & Environment.NewLine))
    End Sub

    <TestMethod>
    Public Sub LogCapture_KeepsLastLinesAndCountsDropped()
        Dim capture As New TaskLogCapture(3)
        For i = 1 To 5
            capture.Add($"l{i}")
        Next
        CollectionAssert.AreEqual(New List(Of String) From {"l3", "l4", "l5"}, capture.GetLines())
        Assert.AreEqual(2, capture.DroppedCount)
    End Sub

    <TestMethod>
    Public Sub LogCapture_TruncatesVeryLongLines()
        Dim capture As New TaskLogCapture()
        capture.Add(New String("x"c, TaskLogCapture.MaxLineLength + 500))
        Assert.IsTrue(capture.GetLines()(0).Length < TaskLogCapture.MaxLineLength + 20)
    End Sub

    <TestMethod>
    Public Sub FormatDuration_ReadableAtEveryScale()
        Assert.AreEqual("850 ms", TaskRunReport.FormatDuration(TimeSpan.FromMilliseconds(850)))
        Assert.AreEqual("12.4 s", TaskRunReport.FormatDuration(TimeSpan.FromMilliseconds(12_430)))
        Assert.AreEqual("4 min 07 s", TaskRunReport.FormatDuration(New TimeSpan(0, 4, 7)))
        Assert.AreEqual("1 h 02 min 05 s", TaskRunReport.FormatDuration(New TimeSpan(1, 2, 5)))
        Assert.AreEqual("0 ms", TaskRunReport.FormatDuration(TimeSpan.FromSeconds(-1)))
    End Sub

    <TestMethod>
    Public Sub LoggerCapture_IsScopedToTheAsyncFlow()
        Dim capture As New TaskLogCapture()
        Logger.CurrentCapture = capture
        Try
            Dim seenInChild As TaskLogCapture = Nothing
            Threading.Tasks.Task.Run(Sub() seenInChild = Logger.CurrentCapture).Wait()
            Assert.AreSame(capture, seenInChild)
        Finally
            Logger.CurrentCapture = Nothing
        End Try
        Assert.IsNull(Logger.CurrentCapture)
    End Sub

End Class
