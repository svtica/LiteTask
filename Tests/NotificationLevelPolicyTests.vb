' -----------------------------------------------------------------------------
' LiteTask - part of LiteSuite. Released under The Unlicense (see LICENSE).
' File:    NotificationLevelPolicyTests.vb
' Author:  LiteTask contributors
' Date:    2026-11-18
' Purpose: Unit tests for LiteTask.NotificationLevelPolicy: an e-mail is sent
'          when the log level of its event is at or above AlertLevel, as a log
'          line is written when its level is at or above the log level.
' -----------------------------------------------------------------------------
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports LiteTask.LiteTask

<TestClass>
Public Class NotificationLevelPolicyTests

    <TestMethod>
    Public Sub DefaultThreshold_SendsFailuresButNotSuccesses()
        Dim threshold = NotificationLevelPolicy.ParseThreshold("Error")
        Assert.IsFalse(NotificationLevelPolicy.ShouldSend(Logger.LogLevel.Info, threshold), "successful run")
        Assert.IsFalse(NotificationLevelPolicy.ShouldSend(Logger.LogLevel.Warning, threshold), "memory warning")
        Assert.IsTrue(NotificationLevelPolicy.ShouldSend(Logger.LogLevel.Error, threshold), "failed run")
        Assert.IsTrue(NotificationLevelPolicy.ShouldSend(Logger.LogLevel.Critical, threshold), "critical memory")
    End Sub

    <TestMethod>
    Public Sub EachThreshold_HasItsOwnEffect()
        Assert.IsTrue(NotificationLevelPolicy.ShouldSend(Logger.LogLevel.Info, Logger.LogLevel.Info))
        Assert.IsTrue(NotificationLevelPolicy.ShouldSend(Logger.LogLevel.Warning, Logger.LogLevel.Warning))
        Assert.IsFalse(NotificationLevelPolicy.ShouldSend(Logger.LogLevel.Info, Logger.LogLevel.Warning))
        Assert.IsFalse(NotificationLevelPolicy.ShouldSend(Logger.LogLevel.Error, Logger.LogLevel.Critical))
    End Sub

    <TestMethod>
    Public Sub ParseThreshold_IsCaseInsensitive()
        Assert.AreEqual(Logger.LogLevel.Warning, NotificationLevelPolicy.ParseThreshold("warning"))
        Assert.AreEqual(Logger.LogLevel.Critical, NotificationLevelPolicy.ParseThreshold(" CRITICAL "))
    End Sub

    <TestMethod>
    Public Sub ParseThreshold_InvalidValue_FallsBackToError()
        Assert.AreEqual(Logger.LogLevel.Error, NotificationLevelPolicy.ParseThreshold(Nothing))
        Assert.AreEqual(Logger.LogLevel.Error, NotificationLevelPolicy.ParseThreshold(""))
        Assert.AreEqual(Logger.LogLevel.Error, NotificationLevelPolicy.ParseThreshold("Verbose"))
        Assert.AreEqual(Logger.LogLevel.Error, NotificationLevelPolicy.ParseThreshold("2"))
    End Sub

    <TestMethod>
    Public Sub PriorityFor_IsHighFromError()
        Assert.AreEqual(NotificationManager.NotificationPriority.Normal, NotificationLevelPolicy.PriorityFor(Logger.LogLevel.Info))
        Assert.AreEqual(NotificationManager.NotificationPriority.Normal, NotificationLevelPolicy.PriorityFor(Logger.LogLevel.Warning))
        Assert.AreEqual(NotificationManager.NotificationPriority.High, NotificationLevelPolicy.PriorityFor(Logger.LogLevel.Error))
        Assert.AreEqual(NotificationManager.NotificationPriority.High, NotificationLevelPolicy.PriorityFor(Logger.LogLevel.Critical))
    End Sub

    <TestMethod>
    Public Sub ParseFlag_IsLenient()
        Assert.IsTrue(NotificationLevelPolicy.ParseFlag("True", False))
        Assert.IsTrue(NotificationLevelPolicy.ParseFlag(" true ", False))
        Assert.IsFalse(NotificationLevelPolicy.ParseFlag("False", True))
        Assert.IsTrue(NotificationLevelPolicy.ParseFlag(Nothing, True))
        Assert.IsFalse(NotificationLevelPolicy.ParseFlag("oui", False))
    End Sub

End Class
