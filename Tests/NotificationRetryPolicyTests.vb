' -----------------------------------------------------------------------------
' LiteTask - part of LiteSuite. Released under The Unlicense (see LICENSE).
' File:    NotificationRetryPolicyTests.vb
' Author:  LiteTask contributors
' Date:    2026-11-09
' Purpose: Unit tests for LiteTask.NotificationRetryPolicy (GLPI #4302): the
'          e-mail retry schedule must spread attempts over minutes, stop after
'          MaxRetries and produce a readable dead-letter file.
' -----------------------------------------------------------------------------
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports LiteTask.LiteTask

<TestClass>
Public Class NotificationRetryPolicyTests

    <TestMethod>
    Public Sub ShouldRetry_StopsAfterMaxRetries()
        Assert.IsTrue(NotificationRetryPolicy.ShouldRetry(0))
        Assert.IsTrue(NotificationRetryPolicy.ShouldRetry(NotificationRetryPolicy.MaxRetries - 1))
        Assert.IsFalse(NotificationRetryPolicy.ShouldRetry(NotificationRetryPolicy.MaxRetries))
    End Sub

    <TestMethod>
    Public Sub GetRetryDelay_IsIncreasingAndMeasuredInMinutes()
        Dim d1 = NotificationRetryPolicy.GetRetryDelay(1)
        Dim d2 = NotificationRetryPolicy.GetRetryDelay(2)
        Dim d3 = NotificationRetryPolicy.GetRetryDelay(3)
        Assert.IsTrue(d1 >= TimeSpan.FromMinutes(1), d1.ToString())
        Assert.IsTrue(d2 > d1 AndAlso d3 > d2)
    End Sub

    <TestMethod>
    Public Sub GetRetryDelay_OutOfRange_IsClamped()
        Assert.AreEqual(NotificationRetryPolicy.GetRetryDelay(1), NotificationRetryPolicy.GetRetryDelay(0))
        Assert.AreEqual(NotificationRetryPolicy.GetRetryDelay(3), NotificationRetryPolicy.GetRetryDelay(10))
    End Sub

    <TestMethod>
    Public Sub TotalRetryWindow_CoversAtLeastTwentyMinutes()
        ' On 2026-10-27 SMTP was back within 30 minutes; immediate retries (v1.2.1.9)
        ' and the 10/20/30 s backoff (v1.2.1.11) both gave up within one minute.
        Assert.IsTrue(NotificationRetryPolicy.TotalRetryWindow() >= TimeSpan.FromMinutes(20))
    End Sub

    <TestMethod>
    Public Sub BuildDeadLetterFileName_IsSortableAndUnique()
        Dim id = Guid.Parse("11111111-2222-3333-4444-555555555555")
        Dim name = NotificationRetryPolicy.BuildDeadLetterFileName(New DateTime(2026, 10, 27, 12, 1, 34), id)
        Assert.AreEqual("20261027_120134_11111111222233334444555555555555.txt", name)
    End Sub

    <TestMethod>
    Public Sub FormatDeadLetter_ContainsSubjectBodyAndError()
        Dim text = NotificationRetryPolicy.FormatDeadLetter("FATAL ERROR: Task Execution Failed - Extracteur",
                                                            "Error: Action failed: Extracteur SQL",
                                                            New DateTime(2026, 10, 27, 12, 1, 29), 4,
                                                            "Failure sending mail. (No such host is known.)")
        Assert.IsTrue(text.Contains("Subject:   FATAL ERROR: Task Execution Failed - Extracteur"), text)
        Assert.IsTrue(text.Contains("Attempts:  4"), text)
        Assert.IsTrue(text.Contains("No such host is known"), text)
        Assert.IsTrue(text.Contains("Error: Action failed: Extracteur SQL"), text)
    End Sub

    <TestMethod>
    Public Sub FormatDeadLetter_NullFields_DoNotThrow()
        Dim text = NotificationRetryPolicy.FormatDeadLetter(Nothing, Nothing, DateTime.MinValue, 1, Nothing)
        Assert.IsTrue(text.Contains("Subject:"))
    End Sub

End Class
