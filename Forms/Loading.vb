Imports System.Windows.Forms

Public Class Loading
    Private Sub Loading_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim timer As New Timer()
        timer.Interval = 3000 ' 3 seconds
        AddHandler timer.Tick, AddressOf CloseLoadingForm
        timer.Start()
    End Sub

    Private Sub CloseLoadingForm(sender As Object, e As EventArgs)
        Dim timer As Timer = DirectCast(sender, Timer)
        timer.Stop()
        timer.Dispose()
        Close()
    End Sub
End Class
