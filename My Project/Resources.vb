Imports System.Drawing


Namespace LiteTask
    Partial Public Class Resources
        Public Shared Function GetEmbeddedImage(name As String) As Image
            Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
            Dim resourceName As String = $"LiteTask.res.Images.{name}"
            Using stream As Stream = assembly.GetManifestResourceStream(resourceName)
                If stream Is Nothing Then
                    Throw New ArgumentException($"Resource not found: {resourceName}")
                End If
                Return Image.FromStream(stream)
            End Using
        End Function

        Public Shared Function GetEmbeddedIcon(name As String) As Icon
            Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
            Dim resourceName As String = $"LiteTask.res.Icons.{name}"
            Using stream As Stream = assembly.GetManifestResourceStream(resourceName)
                If stream Is Nothing Then
                    Throw New ArgumentException($"Resource not found: {resourceName}")
                End If
                Return New Icon(stream)
            End Using
        End Function
    End Class
End Namespace