' -----------------------------------------------------------------------------
' LiteTask - part of LiteSuite. Released under The Unlicense (see LICENSE).
' File:    LiteSuiteCredentialStoreTests.vb
' Author:  LiteTask contributors
' Date:    2026-11-12
' Purpose: Unit tests of LiteSuiteCredentialStore: names, a real round trip in
'          the Windows Credential Manager of the test account (cleaned up), and
'          the same target format as LiteDeploy ("LiteSuite:<name>").
' -----------------------------------------------------------------------------
Imports System.ComponentModel
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports LiteTask.LiteTask

<TestClass>
Public Class LiteSuiteCredentialStoreTests

    <TestMethod>
    Public Sub Names()
        Assert.IsTrue(LiteSuiteCredentialStore.IsValidName("Compte deploiement"))
        Assert.IsTrue(LiteSuiteCredentialStore.IsValidName("svc_deploy@domain"))
        Assert.IsFalse(LiteSuiteCredentialStore.IsValidName(""))
        Assert.IsFalse(LiteSuiteCredentialStore.IsValidName(" x"))
        Assert.IsFalse(LiteSuiteCredentialStore.IsValidName("a:b"))
    End Sub

    <TestMethod>
    Public Sub SameTargetFormatAsLiteDeploy()
        Assert.AreEqual("LiteSuite:", LiteSuiteCredentialStore.Prefix)
    End Sub

    <TestMethod>
    Public Sub SaveReadListDelete_RoundTrip()
        Dim name = "Test " & Guid.NewGuid().ToString("N").Substring(0, 8)
        Try
            LiteSuiteCredentialStore.Save(name, "DOMAIN\user", "p@ss w""ord")
        Catch ex As Win32Exception
            Assert.Inconclusive("Credential Manager not available for this account: " & ex.Message)
        End Try
        Try
            Dim c = LiteSuiteCredentialStore.Read(name)
            Assert.AreEqual("DOMAIN\user", c.UserName)
            Assert.AreEqual("p@ss w""ord", c.Password)
            Assert.IsTrue(LiteSuiteCredentialStore.List().Any(Function(x) x.Name = name AndAlso x.Password Is Nothing))

            ' Through LiteTask's CredentialManager, as the tasks read it
            Dim manager As New CredentialManager(Nothing)
            Assert.IsTrue(manager.GetAllCredentialTargets().Contains($"(LiteSuite, {name})"))
            Dim info = manager.GetCredential($"(LiteSuite, {name})", "")
            Assert.AreEqual("DOMAIN\user", info.Username)
            Assert.AreEqual("p@ss w""ord", info.Password)
        Finally
            Assert.IsTrue(LiteSuiteCredentialStore.Delete(name))
        End Try
        Assert.IsNull(LiteSuiteCredentialStore.Read(name))
        Assert.IsFalse(LiteSuiteCredentialStore.Delete(name))
    End Sub
End Class
