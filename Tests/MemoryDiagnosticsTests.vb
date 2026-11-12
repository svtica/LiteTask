' -----------------------------------------------------------------------------
' LiteTask - part of LiteSuite. Released under The Unlicense (see LICENSE).
' File:    MemoryDiagnosticsTests.vb
' Author:  LiteTask contributors
' Date:    2026-11-11
' Purpose: Unit tests for LiteTask.MemoryDiagnostics: the per-task memory line
'          must show the absolute values and signed deltas, invariant culture.
' -----------------------------------------------------------------------------
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports LiteTask.LiteTask

<TestClass>
Public Class MemoryDiagnosticsTests

    Private Const MB As Long = 1048576

    <TestMethod>
    Public Sub FormatDelta_ShowsValuesAndSignedDeltas()
        Dim before As New MemorySnapshot With {.PrivateBytes = 400 * MB, .GcHeapBytes = 82 * MB, .HandleCount = 1198, .Gen2Collections = 10}
        Dim after As New MemorySnapshot With {.PrivateBytes = 412 * MB + MB \ 2, .GcHeapBytes = 85 * MB, .HandleCount = 1204, .Gen2Collections = 11}
        Assert.AreEqual("private 412.5 MB (+12.5), GC heap 85.0 MB (+3.0), handles 1204 (+6), gen2 GCs +1",
                        MemoryDiagnostics.FormatDelta(before, after))
    End Sub

    <TestMethod>
    Public Sub FormatDelta_NegativeDeltas()
        Dim before As New MemorySnapshot With {.PrivateBytes = 500 * MB, .GcHeapBytes = 90 * MB, .HandleCount = 1300}
        Dim after As New MemorySnapshot With {.PrivateBytes = 450 * MB, .GcHeapBytes = 60 * MB, .HandleCount = 1290}
        Dim text = MemoryDiagnostics.FormatDelta(before, after)
        StringAssert.Contains(text, "private 450.0 MB (-50.0)")
        StringAssert.Contains(text, "handles 1290 (-10)")
    End Sub

    <TestMethod>
    Public Sub Capture_ReadsCurrentProcess()
        Dim snapshot = MemorySnapshot.Capture()
        Assert.IsTrue(snapshot.GcHeapBytes > 0)
        Assert.IsTrue(snapshot.PrivateBytes > 0)
    End Sub

End Class
