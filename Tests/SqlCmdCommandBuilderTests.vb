' -----------------------------------------------------------------------------
' Copyright (c) svtica. All rights reserved.
' File:    SqlCmdCommandBuilderTests.vb
' Author:  LiteTask contributors
' Date:    2026-11-10
' Purpose: Pins the sqlcmd argument string to the exact 1.2.1.12 layout for
'          both authentication modes (GLPI #4302, "no behaviour change").
'          The expected strings below are copied from TaskRunner.vb at
'          commit 540fd27 (v1.2.1.12), lines 504 and 507.
' -----------------------------------------------------------------------------
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports LiteTask.LiteTask

<TestClass>
Public Class SqlCmdCommandBuilderTests

    Private Const Server As String = "SQLHOST"
    Private Const Database As String = "MyDb"
    Private Const TempFile As String = "C:\Apps\LiteTask\LiteTaskData\temp\sqlcmd_20261027120049_0123456789abcdef0123456789abcdef.sql"

    ' v1.2.1.12: sqlcmdArgs = $"""-S"" ""{server}"" ""-d"" ""{database}"" ""-I"" ""-h-1"" ""-i"" ""{tempFile}"" -E"
    <TestMethod>
    Public Sub BuildArguments_WindowsAuthentication_IdenticalTo_1_2_1_12()
        Dim expected = $"""-S"" ""{Server}"" ""-d"" ""{Database}"" ""-I"" ""-h-1"" ""-i"" ""{TempFile}"" -E"
        Assert.AreEqual(expected, SqlCmdCommandBuilder.BuildArguments(Server, Database, TempFile, Nothing))
    End Sub

    ' v1.2.1.12: sqlcmdArgs = $"""-S"" ""{server}"" ""-d"" ""{database}"" ""-I"" ""-h-1"" ""-i"" ""{tempFile}"" -U ""{credential.Username}"""
    <TestMethod>
    Public Sub BuildArguments_SqlAuthentication_IdenticalTo_1_2_1_12()
        Dim expected = $"""-S"" ""{Server}"" ""-d"" ""{Database}"" ""-I"" ""-h-1"" ""-i"" ""{TempFile}"" -U ""svc_extract"""
        Assert.AreEqual(expected, SqlCmdCommandBuilder.BuildArguments(Server, Database, TempFile, "svc_extract"))
    End Sub

    ' 1.2.1.12 emitted -U "" for a credential with an empty user name; keep that too.
    <TestMethod>
    Public Sub BuildArguments_EmptyUsername_StillUsesSqlAuthenticationAsIn_1_2_1_12()
        Dim expected = $"""-S"" ""{Server}"" ""-d"" ""{Database}"" ""-I"" ""-h-1"" ""-i"" ""{TempFile}"" -U """""
        Assert.AreEqual(expected, SqlCmdCommandBuilder.BuildArguments(Server, Database, TempFile, ""))
    End Sub

    <TestMethod>
    Public Sub BuildArguments_NeverContainsPasswordOrNewFlags()
        For Each user In {Nothing, "svc_extract"}
            Dim args = SqlCmdCommandBuilder.BuildArguments(Server, Database, TempFile, user)
            Assert.IsFalse(args.Contains("-P"), "password must only travel through SQLCMDPASSWORD")
            For Each flag In {" -b", " -r", " -l", " -t", " -V"}
                Assert.IsFalse(args.Contains(flag & " ") OrElse args.EndsWith(flag), $"flag '{flag.Trim()}' must not be added in 1.2.1.13")
            Next
        Next
    End Sub

End Class
