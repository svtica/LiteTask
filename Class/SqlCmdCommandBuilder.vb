' -----------------------------------------------------------------------------
' Copyright (c) svtica. All rights reserved.
' File:    SqlCmdCommandBuilder.vb
' Author:  LiteTask contributors
' Date:    2026-11-10
' Purpose: Build the sqlcmd argument list used by TaskRunner for SQL actions.
'          The layout is byte-for-byte the one shipped in 1.2.1.12 and is
'          pinned by SqlCmdCommandBuilderTests: v1.2.1.13 (GLPI #4302) is a
'          diagnostics-only release and must launch exactly the same command
'          lines. Any flag change (-b, -r, -l, -t...) is a separate decision.
'          The password is never on the command line: TaskRunner passes it via
'          the SQLCMDPASSWORD environment variable of the child process.
' -----------------------------------------------------------------------------
Namespace LiteTask
    Public Module SqlCmdCommandBuilder

        ''' <summary>
        ''' sqlcmd arguments exactly as 1.2.1.12 built them. <paramref name="username"/>
        ''' Nothing (no credential) selects integrated security (-E); any other value,
        ''' even empty, is emitted with -U as before, and the caller supplies the
        ''' password through SQLCMDPASSWORD.
        ''' </summary>
        Public Function BuildArguments(server As String, database As String, inputFile As String, username As String) As String
            If username Is Nothing Then
                Return $"""-S"" ""{server}"" ""-d"" ""{database}"" ""-I"" ""-h-1"" ""-i"" ""{inputFile}"" -E"
            End If
            Return $"""-S"" ""{server}"" ""-d"" ""{database}"" ""-I"" ""-h-1"" ""-i"" ""{inputFile}"" -U ""{username}"""
        End Function

    End Module
End Namespace
