' -----------------------------------------------------------------------------
' LiteTask - part of LiteSuite. Released under The Unlicense (see LICENSE).
' File:    AppIcon.vb
' Author:  LiteTask contributors
' Date:    2026-11-11
' Purpose: The LiteTask icon, loaded once and shared by every window and the
'          tray icon. Each form used to embed its own copy of logo.ico (about
'          230 KB, five copies in the .resx files) and decode it on every open.
' -----------------------------------------------------------------------------
Imports System.Drawing
Imports System.Reflection

Namespace LiteTask

    Public NotInheritable Class AppIcon
        Private Sub New()
        End Sub

        Private Shared ReadOnly _icon As New Lazy(Of Icon)(AddressOf Load)

        ''' <summary>The application icon, or Nothing when it cannot be loaded (Windows default icon).</summary>
        Public Shared ReadOnly Property Value As Icon
            Get
                Return _icon.Value
            End Get
        End Property

        Private Shared Function Load() As Icon
            ' 1. Embedded resource (res\ico\logo.ico, EmbeddedResource in LiteTask.vbproj)
            Try
                Dim asm = Assembly.GetExecutingAssembly()
                Dim name = asm.GetManifestResourceNames().
                    FirstOrDefault(Function(n) n.EndsWith("logo.ico", StringComparison.OrdinalIgnoreCase))
                If name IsNot Nothing Then
                    Using stream = asm.GetManifestResourceStream(name)
                        If stream IsNot Nothing Then Return New Icon(stream)
                    End Using
                End If
            Catch
            End Try

            ' 2. File copied next to the executable (res\ico\logo.ico)
            Try
                Dim filePath = Path.Combine(AppContext.BaseDirectory, "res", "ico", "logo.ico")
                If File.Exists(filePath) Then Return New Icon(filePath)
            Catch
            End Try

            ' 3. Icon of the executable itself
            Try
                Return Icon.ExtractAssociatedIcon(Application.ExecutablePath)
            Catch
                Return Nothing
            End Try
        End Function
    End Class

End Namespace
