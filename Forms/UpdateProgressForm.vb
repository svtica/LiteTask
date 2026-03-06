Imports System.Drawing

Namespace LiteTask
    Public Class UpdateProgressForm
        Inherits Form

        Private _logTextBox As RichTextBox
        Private _progressBar As ProgressBar
        Private _closeButton As Button
        Private _titleLabel As Label

        Public Sub New()
            InitializeComponent()
            Me.Translate()
        End Sub

        Private Sub InitializeComponent()
            SuspendLayout()

            ' Title label
            _titleLabel = New Label With {
                .Name = "_titleLabel",
                .Text = "Update Progress",
                .Dock = DockStyle.Top,
                .Height = 30,
                .Font = New Font(Font.FontFamily, 10, FontStyle.Bold),
                .TextAlign = ContentAlignment.MiddleLeft,
                .Padding = New Padding(8, 0, 0, 0)
            }

            ' Progress bar
            _progressBar = New ProgressBar With {
                .Name = "_progressBar",
                .Dock = DockStyle.Top,
                .Height = 24,
                .Style = ProgressBarStyle.Continuous,
                .Minimum = 0,
                .Maximum = 100,
                .Value = 0,
                .Margin = New Padding(10, 4, 10, 4)
            }

            ' Log text box
            _logTextBox = New RichTextBox With {
                .Name = "_logTextBox",
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .BackColor = Color.FromArgb(30, 30, 30),
                .ForeColor = Color.White,
                .Font = New Font("Consolas", 9),
                .BorderStyle = BorderStyle.None,
                .ScrollBars = RichTextBoxScrollBars.Vertical
            }

            ' Close button
            _closeButton = New Button With {
                .Name = "_closeButton",
                .Text = "Close",
                .Dock = DockStyle.Bottom,
                .Height = 35,
                .Enabled = False
            }
            AddHandler _closeButton.Click, Sub(s, e) Close()

            ' Form properties
            Text = "LiteTask - Update"
            Size = New Size(600, 420)
            MinimumSize = New Size(450, 300)
            StartPosition = FormStartPosition.CenterParent
            FormBorderStyle = FormBorderStyle.FixedDialog
            MaximizeBox = False
            MinimizeBox = False
            ' Add controls (order matters for Dock layout)
            Controls.Add(_logTextBox)
            Controls.Add(_progressBar)
            Controls.Add(_titleLabel)
            Controls.Add(_closeButton)

            ResumeLayout(False)
        End Sub

        ''' <summary>
        ''' Appends a status message to the log with the specified color.
        ''' </summary>
        Public Sub AppendLog(message As String, Optional color As Color = Nothing)
            If Me.InvokeRequired Then
                Me.Invoke(Sub() AppendLog(message, color))
                Return
            End If

            If color = Nothing Then color = Color.LightGray

            _logTextBox.SelectionStart = _logTextBox.TextLength
            _logTextBox.SelectionLength = 0
            _logTextBox.SelectionColor = color
            _logTextBox.AppendText(message & Environment.NewLine)
            _logTextBox.ScrollToCaret()
        End Sub

        ''' <summary>
        ''' Updates the progress bar value (0-100).
        ''' </summary>
        Public Sub SetProgress(value As Integer)
            If Me.InvokeRequired Then
                Me.Invoke(Sub() SetProgress(value))
                Return
            End If

            _progressBar.Value = Math.Min(Math.Max(value, 0), 100)
        End Sub

        ''' <summary>
        ''' Marks the update as finished, enabling the close button.
        ''' </summary>
        Public Sub SetCompleted(success As Boolean)
            If Me.InvokeRequired Then
                Me.Invoke(Sub() SetCompleted(success))
                Return
            End If

            _closeButton.Enabled = True
            _progressBar.Value = 100

            If success Then
                AppendLog(Environment.NewLine & TranslationManager.Instance.GetTranslation(
                    "Update.Progress.Complete", "Update ready. The application will now restart."), Color.LimeGreen)
            Else
                AppendLog(Environment.NewLine & TranslationManager.Instance.GetTranslation(
                    "Update.Progress.Failed", "Update failed. Check the log above for details."), Color.OrangeRed)
            End If
        End Sub

        Protected Overrides Sub OnShown(e As EventArgs)
            MyBase.OnShown(e)
            If Owner IsNot Nothing Then Icon = Owner.Icon
        End Sub

        Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
            If Not _closeButton.Enabled Then
                e.Cancel = True
                Return
            End If
            MyBase.OnFormClosing(e)
        End Sub
    End Class
End Namespace
