' -----------------------------------------------------------------------------
' LiteTask - part of LiteSuite. Released under The Unlicense (see LICENSE).
' File:    ProcessFailureFormatter.vb
' Author:  LiteTask contributors
' Date:    2026-11-10
' Purpose: Build the failure message returned by TaskRunner when a child
'          process (sqlcmd, executables) exits with a non-zero code: exit code,
'          duration, stderr and the last lines of stdout. sqlcmd writes its
'          error messages on stdout unless -r is given, and the previous text
'          ("Error: Process exited with code 1") lost them (GLPI #4302).
'          The message always starts with "Error:" so the success/failure
'          decision of the callers (StartsWith("Error:")) is unchanged.
' -----------------------------------------------------------------------------
Namespace LiteTask
    Public Module ProcessFailureFormatter

        Public Const DefaultTailLines As Integer = 40
        Public Const DefaultTailChars As Integer = 8000

        Public Function BuildFailureMessage(exitCode As Integer, standardOutput As String, standardError As String,
                                            duration As TimeSpan,
                                            Optional maxTailLines As Integer = DefaultTailLines,
                                            Optional maxTailChars As Integer = DefaultTailChars) As String
            Dim sb As New StringBuilder()
            sb.Append($"Error: Process exited with code {exitCode} after {duration.TotalSeconds:F1} s")

            Dim err = If(standardError, String.Empty).Trim()
            Dim out = If(standardOutput, String.Empty).Trim()

            If err.Length > 0 Then
                sb.AppendLine()
                sb.Append("stderr: ")
                sb.Append(TailChars(err, maxTailChars))
            End If

            If out.Length > 0 Then
                Dim tail = TailLines(out, maxTailLines)
                sb.AppendLine()
                sb.Append($"stdout (last {CountLines(tail)} of {CountLines(out)} lines): ")
                sb.Append(TailChars(tail, maxTailChars))
            End If

            Return sb.ToString()
        End Function

        ''' <summary>Last <paramref name="maxLines"/> lines of the text (any line ending).</summary>
        Public Function TailLines(text As String, maxLines As Integer) As String
            If String.IsNullOrEmpty(text) Then Return String.Empty
            Dim lines = text.Replace(vbCrLf, vbLf).Split(vbLf(0))
            If maxLines <= 0 OrElse lines.Length <= maxLines Then Return text
            Return "..." & Environment.NewLine & String.Join(Environment.NewLine, lines, lines.Length - maxLines, maxLines)
        End Function

        ''' <summary>Last <paramref name="maxChars"/> characters, prefixed with "..." when truncated.</summary>
        Public Function TailChars(text As String, maxChars As Integer) As String
            If text Is Nothing Then Return String.Empty
            If maxChars <= 0 OrElse text.Length <= maxChars Then Return text
            Return "..." & text.Substring(text.Length - maxChars)
        End Function

        Public Function CountLines(text As String) As Integer
            If String.IsNullOrEmpty(text) Then Return 0
            Return text.Replace(vbCrLf, vbLf).Split(vbLf(0)).Length
        End Function

    End Module
End Namespace
