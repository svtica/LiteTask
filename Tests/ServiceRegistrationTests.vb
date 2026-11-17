' -----------------------------------------------------------------------------
' LiteTask - part of LiteSuite. Released under The Unlicense (see LICENSE).
' File:    ServiceRegistrationTests.vb
' Author:  LiteTask contributors
' Date:    2026-11-17
' Purpose: Pins the sc.exe command line that registers LiteTaskService (menu
'          and -register): the service must run "<LiteTask.exe>" -service,
'          quoted, so that the self-updater finds it by its ImagePath.
' -----------------------------------------------------------------------------
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports LiteTask.LiteTask

<TestClass>
Public Class ServiceRegistrationTests

    Private Const ExeWithSpaces As String = "C:\Program Files\LiteTask\LiteTask.exe"
    Private Const ExeWithoutSpaces As String = "C:\Apps\LiteTask\LiteTask.exe"

    <TestMethod>
    Public Sub BuildImagePath_QuotesTheExecutableAndAddsServiceArgument()
        Assert.AreEqual("""C:\Program Files\LiteTask\LiteTask.exe"" -service",
                        ServiceRegistration.BuildImagePath(ExeWithSpaces))
    End Sub

    ' Expected text, as typed at a prompt:
    ' create LiteTaskService binPath= "\"C:\Program Files\LiteTask\LiteTask.exe\" -service" start= auto
    <TestMethod>
    Public Sub BuildCreateArguments_PathWithSpaces_EscapesInnerQuotes()
        Dim expected = "create LiteTaskService binPath= ""\""C:\Program Files\LiteTask\LiteTask.exe\"" -service"" start= auto"
        Assert.AreEqual(expected, ServiceRegistration.BuildCreateArguments(ExeWithSpaces))
    End Sub

    <TestMethod>
    Public Sub BuildCreateArguments_PathWithoutSpaces_IsQuotedToo()
        Dim expected = "create LiteTaskService binPath= ""\""C:\Apps\LiteTask\LiteTask.exe\"" -service"" start= auto"
        Assert.AreEqual(expected, ServiceRegistration.BuildCreateArguments(ExeWithoutSpaces))
    End Sub

    <TestMethod>
    Public Sub BuildCreateArguments_RunsTheExecutableNotTheDll()
        Dim args = ServiceRegistration.BuildCreateArguments(ExeWithSpaces)
        StringAssert.Contains(args, "LiteTask.exe\"" -service")
        Assert.IsFalse(args.Contains(".dll"), "The service must never be registered on LiteTask.dll.")
    End Sub

    <TestMethod>
    Public Sub BuildDeleteArguments_DeletesLiteTaskService()
        Assert.AreEqual("delete LiteTaskService", ServiceRegistration.BuildDeleteArguments())
    End Sub

End Class
