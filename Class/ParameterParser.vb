' -----------------------------------------------------------------------------
' Copyright (c) svtica. All rights reserved.
' File:    ParameterParser.vb
' Author:  LiteTask contributors
' Date:    2026-04-25
' Purpose: Parse parameter strings used by TaskRunner. Supports three forms
'          that can be mixed in the same string:
'            key=value
'            key="value with spaces"
'            -Name Value                  (PowerShell-style)
'            -Name "Value with spaces"    (PowerShell-style)
'            -Switch                      (no value -> stored as Nothing)
' -----------------------------------------------------------------------------
Namespace LiteTask
    Public Module ParameterParser

        ''' <summary>
        ''' Parses a parameter string. Supports both `key=value` and
        ''' PowerShell-style `-Name Value` / `-Switch` syntax in the same
        ''' input. Returns an ordered dictionary mapping each parameter name
        ''' to its value. A switch parameter (no value provided) is stored
        ''' as Nothing so callers can distinguish it from an empty string.
        ''' Tokens without an `=` and that don't start with `-Letter` are
        ''' skipped to preserve the original lenient behavior.
        ''' Last occurrence wins for duplicate keys.
        ''' </summary>
        Public Function Parse(parameters As String) As Dictionary(Of String, Object)
            Dim result As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            If String.IsNullOrEmpty(parameters) Then Return result

            Dim i As Integer = 0
            Dim len As Integer = parameters.Length

            While i < len
                ' Skip leading whitespace
                While i < len AndAlso Char.IsWhiteSpace(parameters(i))
                    i += 1
                End While
                If i >= len Then Exit While

                ' PowerShell-style: -Name [Value]  (a dash followed by a letter)
                If parameters(i) = "-"c AndAlso i + 1 < len AndAlso IsNameStart(parameters(i + 1)) Then
                    i += 1 ' consume '-'
                    Dim keyStart As Integer = i
                    While i < len AndAlso Not Char.IsWhiteSpace(parameters(i)) AndAlso parameters(i) <> "="c
                        i += 1
                    End While
                    Dim psKey As String = parameters.Substring(keyStart, i - keyStart)

                    ' Allow `-Name=Value` as a convenience
                    If i < len AndAlso parameters(i) = "="c Then
                        i += 1
                        result(psKey) = ReadValue(parameters, i)
                        Continue While
                    End If

                    ' Skip whitespace between name and value
                    Dim valueScan As Integer = i
                    While valueScan < len AndAlso Char.IsWhiteSpace(parameters(valueScan))
                        valueScan += 1
                    End While

                    ' If next non-whitespace is another -Name, this is a switch.
                    If valueScan >= len OrElse
                       (parameters(valueScan) = "-"c AndAlso valueScan + 1 < len AndAlso IsNameStart(parameters(valueScan + 1))) Then
                        result(psKey) = Nothing
                        ' Don't consume the next token; loop will pick it up.
                    Else
                        i = valueScan
                        result(psKey) = ReadValue(parameters, i)
                    End If
                    Continue While
                End If

                ' key=value form
                Dim keyStartEq As Integer = i
                While i < len AndAlso parameters(i) <> "="c AndAlso Not Char.IsWhiteSpace(parameters(i))
                    i += 1
                End While

                If i >= len OrElse parameters(i) <> "="c Then
                    ' Bare token without '=': skip to preserve original lenient behavior.
                    Continue While
                End If

                Dim key As String = parameters.Substring(keyStartEq, i - keyStartEq)
                i += 1 ' consume '='

                If String.IsNullOrEmpty(key) Then
                    SkipValue(parameters, i)
                    Continue While
                End If

                result(key) = ReadValue(parameters, i)
            End While

            Return result
        End Function

        Private Function IsNameStart(c As Char) As Boolean
            Return Char.IsLetter(c) OrElse c = "_"c
        End Function

        Private Function ReadValue(s As String, ByRef i As Integer) As String
            Dim len As Integer = s.Length
            If i >= len Then Return String.Empty

            If s(i) = """"c Then
                i += 1
                Dim valueStart As Integer = i
                While i < len AndAlso s(i) <> """"c
                    i += 1
                End While
                Dim value As String = s.Substring(valueStart, i - valueStart)
                If i < len Then i += 1 ' consume closing quote
                Return value
            End If

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
