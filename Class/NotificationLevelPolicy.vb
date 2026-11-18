' -----------------------------------------------------------------------------
' LiteTask - part of LiteSuite. Released under The Unlicense (see LICENSE).
' File:    NotificationLevelPolicy.vb
' Author:  LiteTask contributors
' Date:    2026-11-18
' Purpose: Decides whether an e-mail notification is sent, from the log level
'          of the event it reports and EmailSettings/AlertLevel. Each
'          notification carries the same level as the log line of its event, so
'          the threshold works like the log level: a notification is sent when
'          its level is at or above the threshold.
'          Before, the level was derived from the mail priority (Normal = Error,
'          High = Critical): a successful run was treated as an Error, so the
'          default threshold e-mailed every run and only "Critical" changed
'          anything. Kept out of NotificationManager so it can be unit-tested.
' -----------------------------------------------------------------------------
Imports LiteTask.LiteTask.NotificationManager

Namespace LiteTask
    Public NotInheritable Class NotificationLevelPolicy

        ''' <summary>Threshold used when AlertLevel is missing or not a log level name.</summary>
        Public Const DefaultThreshold As Logger.LogLevel = Logger.LogLevel.Error

        Private Sub New()
        End Sub

        ''' <summary>
        ''' Parses EmailSettings/AlertLevel (Debug, Info, Warning, Error, Critical,
        ''' case-insensitive). Anything else gives <see cref="DefaultThreshold"/>.
        ''' </summary>
        Public Shared Function ParseThreshold(value As String) As Logger.LogLevel
            Dim parsed As Logger.LogLevel
            If Not String.IsNullOrWhiteSpace(value) AndAlso
               [Enum].TryParse(value.Trim(), True, parsed) AndAlso
               [Enum].IsDefined(GetType(Logger.LogLevel), parsed) AndAlso
               Not IsNumeric(value.Trim()) Then
                Return parsed
            End If
            Return DefaultThreshold
        End Function

        ''' <summary>True when a notification of this level passes the threshold.</summary>
        Public Shared Function ShouldSend(level As Logger.LogLevel, threshold As Logger.LogLevel) As Boolean
            Return level >= threshold
        End Function

        ''' <summary>Mail priority: High for Error and Critical, Normal below.</summary>
        Public Shared Function PriorityFor(level As Logger.LogLevel) As NotificationPriority
            Return If(level >= Logger.LogLevel.Error, NotificationPriority.High, NotificationPriority.Normal)
        End Function

        ''' <summary>Lenient Boolean parse of a settings value (missing or invalid = defaultValue).</summary>
        Public Shared Function ParseFlag(value As String, defaultValue As Boolean) As Boolean
            Dim parsed As Boolean
            Return If(Boolean.TryParse(If(value, "").Trim(), parsed), parsed, defaultValue)
        End Function

    End Class
End Namespace
