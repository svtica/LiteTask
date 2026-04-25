' -----------------------------------------------------------------------------
' Copyright (c) svtica. All rights reserved.
' File:    ParameterParser.vb
' Author:  LiteTask contributors
' Date:    2026-04-25
' Purpose: Parse "key=value" parameter strings used by TaskRunner. Supports
'          cmd-style quoting (key="value with spaces") while remaining
'          backward-compatible with the original whitespace-delimited form.
' -----------------------------------------------------------------------------
Namespace LiteTask
    Public Module ParameterParser

        ''' <summary>
        ''' Parses a parameter string of the form `key=value [key2=value2 ...]`.
        ''' Values may be quoted with double quotes to include spaces:
        '''   key="value with spaces"
        ''' Tokens without an `=` are skipped, as are `=value` fragments.
        ''' Last occurrence wins for duplicate keys.
        ''' </summary>
        Public Function Parse(parameters As String) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)
            If String.IsNullOrEmpty(parameters) Then Return result

            Dim i As Integer = 0
            Dim len As Integer = parameters.Length

            While i < len
                ' Skip leading whitespace
                While i < len AndAlso Char.IsWhiteSpace(parameters(i))
                    i += 1
                End While
                If i >= len Then Exit While

                ' Read key: up to '=' or whitespace
                Dim keyStart As Integer = i
                While i < len AndAlso parameters(i) <> "="c AndAlso Not Char.IsWhiteSpace(parameters(i))
                    i += 1
                End While

                ' No '=' for this token: skip token entirely (preserves original behavior
                ' where bare words without '=' were dropped).
                If i >= len OrElse parameters(i) <> "="c Then
                    Continue While
                End If

                Dim key As String = parameters.Substring(keyStart, i - keyStart)
                i += 1 ' consume '='

                ' Empty key (`=value`) is skipped to match original behavior.
                If String.IsNullOrEmpty(key) Then
                    ' Advance past the value so we don't reparse it as a new token
                    SkipValue(parameters, i)
                    Continue While
                End If

                Dim value As String = ReadValue(parameters, i)
                result(key) = value
            End While

            Return result
        End Function

        Private Function ReadValue(s As String, ByRef i As Integer) As String
            Dim len As Integer = s.Length
            If i >= len Then Return String.Empty

            If s(i) = """"c Then
                ' Quoted value: read until next '"' or end of string.
                i += 1
                Dim valueStart As Integer = i
                While i < len AndAlso s(i) <> """"c
                    i += 1
                End While
                Dim value As String = s.Substring(valueStart, i - valueStart)
                If i < len Then i += 1 ' consume closing quote
                Return value
            End If

            ' Unquoted value: read until whitespace.
            Dim unquotedStart As Integer = i
            While i < len AndAlso Not Char.IsWhiteSpace(s(i))
                i += 1
            End While
            Return s.Substring(unquotedStart, i - unquotedStart)
        End Function

        Private Sub SkipValue(s As String, ByRef i As Integer)
            ReadValue(s, i)
        End Sub

    End Module
End Namespace
