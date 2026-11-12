' -----------------------------------------------------------------------------
' LiteTask - part of LiteSuite. Released under The Unlicense (see LICENSE).
' File:    ProcessFailureFormatterTests.vb
' Author:  LiteTask contributors
' Date:    2026-11-10
' Purpose: Unit tests for LiteTask.ProcessFailureFormatter (GLPI #4302): a
'          failed child process must report exit code, duration, stderr and
'          the last lines of stdout, while keeping the "Error:" prefix the
'          callers use to decide success or failure.
' -----------------------------------------------------------------------------
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports System.Linq
Imports LiteTask.LiteTask

<TestClass>
Public Class ProcessFailureFormatterTests

    Private ReadOnly _duration As TimeSpan = TimeSpan.FromSeconds(39.8)

    <TestMethod>
    Public Sub BuildFailureMessage_NoOutput_HasExitCodeAndDurationOnly()
        Dim msg = ProcessFailureFormatter.BuildFailureMessage(1, "", "", _duration)
        Assert.AreEqual("Error: Process exited with code 1 after 39.8 s", msg)
    End Sub

    <TestMethod>
    Public Sub BuildFailureMessage_AlwaysStartsWithErrorPrefix()
        Dim msg = ProcessFailureFormatter.BuildFailureMessage(3, "some output", "some error", _duration)
        Assert.IsTrue(msg.StartsWith("Error:", StringComparison.OrdinalIgnoreCase))
    End Sub

    <TestMethod>
    Public Sub BuildFailureMessage_SqlcmdErrorOnStdout_IsIncluded()
        ' sqlcmd without -r prints its errors on stdout; the exit code alone was all we kept.
        Dim stdout = "Msg 547, Level 16, State 0, Server SQLHOST, Line 12" & vbCrLf &
                     "The INSERT statement conflicted with the FOREIGN KEY constraint"
        Dim msg = ProcessFailureFormatter.BuildFailureMessage(1, stdout, "", _duration)
        Assert.IsTrue(msg.Contains("Msg 547"), msg)
        Assert.IsTrue(msg.Contains("stdout (last 2 of 2 lines):"), msg)
        Assert.IsFalse(msg.Contains("stderr:"), msg)
    End Sub

    <TestMethod>
    Public Sub BuildFailureMessage_StderrAndStdout_BothPresentInOrder()
        Dim msg = ProcessFailureFormatter.BuildFailureMessage(1, "out", "err", _duration)
        Dim errPos = msg.IndexOf("stderr: err", StringComparison.Ordinal)
        Dim outPos = msg.IndexOf("stdout (last 1 of 1 lines): out", StringComparison.Ordinal)
        Assert.IsTrue(errPos > 0 AndAlso outPos > errPos, msg)
    End Sub

    <TestMethod>
    Public Sub BuildFailureMessage_LongStdout_KeepsLastFortyLines()
        Dim lines = Enumerable.Range(1, 100).Select(Function(i) $"line {i}")
        Dim stdout = String.Join(vbCrLf, lines)
        Dim msg = ProcessFailureFormatter.BuildFailureMessage(1, stdout, "", _duration)
        Assert.IsTrue(msg.Contains("stdout (last 40 of 100 lines):"), msg)
        Assert.IsFalse(msg.Contains("line 60" & Environment.NewLine), msg)
        Assert.IsTrue(msg.Contains("line 61"), msg)
        Assert.IsTrue(msg.EndsWith("line 100"), msg)
    End Sub

    <TestMethod>
    Public Sub TailLines_AndTailChars_EdgeCases()
        Assert.AreEqual(String.Empty, ProcessFailureFormatter.TailLines(Nothing, 5))
        Assert.AreEqual("a" & vbLf & "b", ProcessFailureFormatter.TailLines("a" & vbLf & "b", 5))
        Assert.AreEqual(String.Empty, ProcessFailureFormatter.TailChars(Nothing, 5))
        Assert.AreEqual("abc", ProcessFailureFormatter.TailChars("abc", 0))
        Assert.AreEqual("...TAIL", ProcessFailureFormatter.TailChars("aaaaTAIL", 4))
    End Sub

End Class
