' -----------------------------------------------------------------------------
' Copyright (c) svtica. All rights reserved.
' File:    PsExecCommandBuilder.vb
' Author:  LiteTask contributors
' Date:    2026-10-28
' Purpose: Build the PsExec command line used by TaskRunner to run Batch and
'          Executable actions under a stored credential. Pure string logic,
'          kept out of TaskRunner so it can be unit-tested without Windows,
'          PsExec or a real credential.
'
'          Why the password is on the PsExec command line: PsExec has no
'          environment-variable or stdin option for -p. When -p is omitted it
'          prompts on the console (via CONIN$, not stdin), which hangs or fails
'          under the service. The previous implementation passed the literal
'          `-p "%VAR%"` which nothing expanded (UseShellExecute = False), so
'          authentication always failed. A process command line is readable
'          only by the same user and by administrators, so under the service
'          (LocalSystem) it is only visible to local administrators.
' -----------------------------------------------------------------------------
Namespace LiteTask
    Public Module PsExecCommandBuilder

        ''' <summary>
        ''' Builds the full PsExec argument string:
        ''' <c>-accepteula -nobanner [-i] [-h] -u user -p "password" command</c>.
        ''' <paramref name="command"/> is appended verbatim; callers are
        ''' responsible for quoting the program path (see
        ''' <see cref="BuildExecutableCommand"/> / <see cref="BuildBatchCommand"/>).
        ''' </summary>
        Public Function BuildArguments(username As String,
                                       password As String,
                                       requiresElevation As Boolean,
                                       interactive As Boolean,
                                       command As String) As String
            If String.IsNullOrWhiteSpace(username) Then
                Throw New ArgumentException("A username is required to run through PsExec.", NameOf(username))
            End If
            If String.IsNullOrWhiteSpace(command) Then
                Throw New ArgumentException("A command is required to run through PsExec.", NameOf(command))
            End If

            Dim args As New System.Text.StringBuilder("-accepteula -nobanner")
            ' -i (interactive) is only valid in a user session: under the service there
            ' is no visible desktop, so an interactive child can block invisibly.
            If interactive Then args.Append(" -i")
            If requiresElevation Then args.Append(" -h")
            args.Append(" -u ").Append(QuoteArgument(username.Trim()))
            ' Always quote the password so spaces and shell metacharacters survive
            ' PsExec's argv parsing. An empty password is passed as "" so PsExec
            ' does not prompt on the console.
            args.Append(" -p ").Append(QuoteArgument(If(password, String.Empty), forceQuotes:=True))
            args.Append(" ").Append(command.Trim())
            Return args.ToString()
        End Function

        ''' <summary>
        ''' Command for an Executable action. Mirrors the non-credential path:
        ''' an empty target means "run the parameters through cmd.exe /c".
        ''' </summary>
        Public Function BuildExecutableCommand(target As String, parameters As String) As String
            Dim params As String = If(parameters, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(target) Then
                Return $"cmd.exe /c {params}".TrimEnd()
            End If
            Return $"{QuoteArgument(target.Trim(), forceQuotes:=True)} {params}".TrimEnd()
        End Function

        ''' <summary>
        ''' Command for a Batch action: <c>cmd.exe /c "script.bat" args</c> so the
        ''' script receives its parameters as %1..%*.
        ''' </summary>
        Public Function BuildBatchCommand(batchFile As String, parameters As String) As String
            If String.IsNullOrWhiteSpace(batchFile) Then
                Throw New ArgumentException("A batch file path is required.", NameOf(batchFile))
            End If
            Dim params As String = If(parameters, String.Empty).Trim()
            Return $"cmd.exe /c {QuoteArgument(batchFile.Trim(), forceQuotes:=True)} {params}".TrimEnd()
        End Function

        ''' <summary>
        ''' Quotes a single argument using the Windows C runtime / CommandLineToArgvW
        ''' rules (backslashes are only special when they precede a quote), so the
        ''' child receives exactly <paramref name="value"/> in its argv entry.
        ''' Without <paramref name="forceQuotes"/>, values that need no quoting
        ''' are returned unchanged.
        ''' </summary>
        Public Function QuoteArgument(value As String, Optional forceQuotes As Boolean = False) As String
            If value Is Nothing Then value = String.Empty

            Dim needsQuotes As Boolean = forceQuotes OrElse value.Length = 0 OrElse
                                         value.IndexOfAny(New Char() {" "c, Microsoft.VisualBasic.ControlChars.Tab, Microsoft.VisualBasic.ControlChars.Cr, Microsoft.VisualBasic.ControlChars.Lf, """"c}) >= 0
            If Not needsQuotes Then Return value

            Dim sb As New System.Text.StringBuilder()
            sb.Append(""""c)
            Dim i As Integer = 0
            While i < value.Length
                Dim backslashes As Integer = 0
                While i < value.Length AndAlso value(i) = "\"c
                    backslashes += 1
                    i += 1
                End While

                If i = value.Length Then
                    ' Trailing backslashes: double them so the closing quote stays a delimiter
                    sb.Append("\"c, backslashes * 2)
                ElseIf value(i) = """"c Then
                    ' Backslashes before a quote: double them, then escape the quote
                    sb.Append("\"c, backslashes * 2 + 1)
                    sb.Append(""""c)
                    i += 1
                Else
                    sb.Append("\"c, backslashes)
                    sb.Append(value(i))
                    i += 1
                End If
            End While
            sb.Append(""""c)
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' True when <paramref name="line"/> is one of PsExec's own progress messages,
        ''' which it writes to stderr even with <c>-nobanner</c>:
        ''' "Connecting to local system...", "Starting PSEXESVC service on local system...",
        ''' "Copying authentication key to HOST...", "Connecting with PsExec service on HOST...",
        ''' "Starting C:\app.exe on HOST..." and "C:\app.exe exited on HOST with error code 0.".
        ''' Real failures ("Access is denied.", "Couldn't access HOST:", "The user name or
        ''' password is incorrect.") do not match and stay classified as errors.
        ''' </summary>
        Public Function IsStatusLine(line As String) As Boolean
            If String.IsNullOrWhiteSpace(line) Then Return False
            Return StatusLinePattern.IsMatch(line.Trim())
        End Function

        Private ReadOnly StatusLinePattern As New System.Text.RegularExpressions.Regex(
            "^(Connecting (to|with) .+\.\.\.|Starting .+ on .+\.\.\.|Copying .+ to .+\.\.\.|.+ exited on .+ with error code -?\d+\.?)$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.Compiled)

    End Module
End Namespace
