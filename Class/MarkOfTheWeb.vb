' -----------------------------------------------------------------------------
' Copyright (c) svtica. All rights reserved.
' File:    MarkOfTheWeb.vb
' Author:  LiteTask contributors
' Date:    2026-11-11
' Purpose: Removes the Mark of the Web (the NTFS "Zone.Identifier" alternate
'          data stream Windows adds to files that come from the Internet) from
'          the tools LiteTask downloads or finds in its folders. A marked
'          executable started through the shell (LitePM) shows the "Open File -
'          Security Warning" prompt, and marked scripts are refused by the
'          RemoteSigned execution policy. Same effect as PowerShell Unblock-File.
' -----------------------------------------------------------------------------
Namespace LiteTask

    Public NotInheritable Class MarkOfTheWeb
        Private Sub New()
        End Sub

        Private Const StreamSuffix As String = ":Zone.Identifier"

        ''' <summary>True when the file carries a Mark of the Web.</summary>
        Public Shared Function HasMark(filePath As String) As Boolean
            Try
                Return File.Exists(filePath & StreamSuffix)
            Catch
                Return False
            End Try
        End Function

        ''' <summary>Removes the mark from one file. Returns True when a mark was removed.</summary>
        Public Shared Function RemoveFromFile(filePath As String) As Boolean
            Try
                If Not HasMark(filePath) Then Return False
                File.Delete(filePath & StreamSuffix)
                Return True
            Catch
                ' Non-NTFS volume, access denied...: nothing to do
                Return False
            End Try
        End Function

        ''' <summary>Removes the mark from every file under a folder. Returns the number of files unblocked.</summary>
        Public Shared Function RemoveFromDirectory(folder As String, Optional recursive As Boolean = True) As Integer
            Dim count = 0
            Try
                If Not Directory.Exists(folder) Then Return 0
                Dim scope = If(recursive, IO.SearchOption.AllDirectories, IO.SearchOption.TopDirectoryOnly)
                For Each filePath In Directory.EnumerateFiles(folder, "*", scope)
                    If RemoveFromFile(filePath) Then count += 1
                Next
            Catch
            End Try
            Return count
        End Function
    End Class

End Namespace
