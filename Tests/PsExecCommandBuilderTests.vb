' -----------------------------------------------------------------------------
' Copyright (c) svtica. All rights reserved.
' File:    PsExecCommandBuilderTests.vb
' Author:  LiteTask contributors
' Date:    2026-10-28
' Purpose: Unit tests for LiteTask.PsExecCommandBuilder: PsExec argument
'          layout, password quoting and the Batch/Executable command shapes.
' -----------------------------------------------------------------------------
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports LiteTask.LiteTask

<TestClass>
Public Class PsExecCommandBuilderTests

    <TestMethod>
    Public Sub BuildArguments_ServiceContext_NoInteractiveFlag()
        Dim args = PsExecCommandBuilder.BuildArguments("DOMAIN\svc", "p@ss", False, False, "whoami.exe")
        Assert.AreEqual("-accepteula -nobanner -u DOMAIN\svc -p ""p@ss"" whoami.exe", args)
    End Sub

    <TestMethod>
    Public Sub BuildArguments_InteractiveAndElevated_AddsFlagsInOrder()
        Dim args = PsExecCommandBuilder.BuildArguments("user", "pw", True, True, "whoami.exe")
        Assert.AreEqual("-accepteula -nobanner -i -h -u user -p ""pw"" whoami.exe", args)
    End Sub

    <TestMethod>
    Public Sub BuildArguments_PasswordWithSpacesAndQuotes_IsEscaped()
        Dim args = PsExecCommandBuilder.BuildArguments("user", "my ""secret"" pw", False, False, "whoami.exe")
        Assert.IsTrue(args.Contains(" -p ""my \""secret\"" pw"" "), args)
    End Sub

    <TestMethod>
    Public Sub BuildArguments_PasswordEndingWithBackslash_DoublesTrailingBackslash()
        Dim args = PsExecCommandBuilder.BuildArguments("user", "abc\", False, False, "whoami.exe")
        Assert.IsTrue(args.Contains(" -p ""abc\\"" "), args)
    End Sub

    <TestMethod>
    Public Sub BuildArguments_EmptyPassword_PassesEmptyQuotedValue()
        Dim args = PsExecCommandBuilder.BuildArguments("user", "", False, False, "whoami.exe")
        Assert.IsTrue(args.Contains(" -p """" "), args)
    End Sub

    <TestMethod>
    Public Sub BuildArguments_NeverContainsEnvironmentPlaceholder()
        Dim args = PsExecCommandBuilder.BuildArguments("user", "pw", False, False, "whoami.exe")
        Assert.IsFalse(args.Contains("%"), args)
    End Sub

    <TestMethod>
    Public Sub BuildArguments_MissingUsername_Throws()
        Assert.ThrowsException(Of ArgumentException)(
            Sub() PsExecCommandBuilder.BuildArguments("", "pw", False, False, "whoami.exe"))
    End Sub

    <TestMethod>
    Public Sub BuildExecutableCommand_TargetWithSpaces_IsQuoted()
        Dim cmd = PsExecCommandBuilder.BuildExecutableCommand("C:\Program Files\App\tool.exe", "-a 1")
        Assert.AreEqual("""C:\Program Files\App\tool.exe"" -a 1", cmd)
    End Sub

    <TestMethod>
    Public Sub BuildExecutableCommand_NoParameters_NoTrailingSpace()
        Dim cmd = PsExecCommandBuilder.BuildExecutableCommand("whoami.exe", Nothing)
        Assert.AreEqual("""whoami.exe""", cmd)
    End Sub

    <TestMethod>
    Public Sub BuildExecutableCommand_EmptyTarget_RunsParametersThroughCmd()
        Dim cmd = PsExecCommandBuilder.BuildExecutableCommand("", "echo hello")
        Assert.AreEqual("cmd.exe /c echo hello", cmd)
    End Sub

    <TestMethod>
    Public Sub BuildBatchCommand_PassesScriptAndArgs()
        Dim cmd = PsExecCommandBuilder.BuildBatchCommand("C:\Temp\batch_1.bat", "arg1 ""arg two""")
        Assert.AreEqual("cmd.exe /c ""C:\Temp\batch_1.bat"" arg1 ""arg two""", cmd)
    End Sub

    <TestMethod>
    Public Sub QuoteArgument_PlainValue_Unchanged()
        Assert.AreEqual("simple", PsExecCommandBuilder.QuoteArgument("simple"))
    End Sub

    <TestMethod>
    Public Sub QuoteArgument_BackslashesBeforeQuote_AreDoubled()
        ' a\"b must round-trip through argv parsing as a\"b
        Assert.AreEqual("""a\\\""b""", PsExecCommandBuilder.QuoteArgument("a\""b"))
    End Sub

    <TestMethod>
    Public Sub QuoteArgument_InternalBackslashes_LeftAlone()
        Assert.AreEqual("""C:\dir\file.exe""", PsExecCommandBuilder.QuoteArgument("C:\dir\file.exe", forceQuotes:=True))
    End Sub

    <TestMethod>
    Public Sub IsStatusLine_PsExecProgressMessages_AreStatus()
        Dim lines = {
            "Connecting to local system...",
            "Starting PSEXESVC service on local system...",
            "Copying authentication key to XCONCENTMDS...",
            "Connecting with PsExec service on XCONCENTMDS...",
            "Starting C:\SIVT II\Extracteur\ExtracteurVBNET-DTS-CAPII.exe on XCONCENTMDS...",
            "C:\SIVT II\GenRTP Automatique\GenRTP_CAP_II.exe exited on XCONCENTMDS with error code 0.",
            "cmd.exe exited on XCONCENTMDS with error code -1073741510."
        }
        For Each line In lines
            Assert.IsTrue(PsExecCommandBuilder.IsStatusLine(line), line)
        Next
    End Sub

    <TestMethod>
    Public Sub IsStatusLine_RealErrorsAndBlankLines_AreNotStatus()
        Dim lines = {
            "",
            "   ",
            "Access is denied.",
            "Couldn't access XCONCENTMDS:",
            "The user name or password is incorrect.",
            "PsExec could not start cmd.exe on XCONCENTMDS:",
            "Starting import failed"
        }
        For Each line In lines
            Assert.IsFalse(PsExecCommandBuilder.IsStatusLine(line), $"'{line}'")
        Next
        Assert.IsFalse(PsExecCommandBuilder.IsStatusLine(Nothing))
    End Sub

End Class
