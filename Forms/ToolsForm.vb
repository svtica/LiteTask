' -----------------------------------------------------------------------------
' Copyright (c) svtica. All rights reserved.
' File:    ToolsForm.vb
' Author:  LiteTask contributors
' Date:    2026-11-11
' Purpose: Tools > Manage Tools: choose which optional tools (sqlcmd, LitePM,
'          LiteDeploy) this installation uses, download or update them from
'          their GitHub releases, or remove them.
' -----------------------------------------------------------------------------
Imports System.Drawing

Namespace LiteTask
    Public Class ToolsForm
        Inherits Form

        Private ReadOnly _toolManager As ToolManager
        Private _toolsListView As ListView
        Private _logTextBox As TextBox
        Private _installButton As Button
        Private _removeButton As Button
        Private _closeButton As Button
        Private _hintLabel As Label
        Private _loading As Boolean
        Private _busy As Boolean

        Public Sub New(toolManager As ToolManager)
            _toolManager = toolManager
            InitializeComponent()
            LoadTools()
        End Sub

        Private Shared Function Tr(key As String, defaultValue As String) As String
            Return TranslationManager.Instance.GetTranslation(key, defaultValue)
        End Function

        Private Sub InitializeComponent()
            SuspendLayout()

            _hintLabel = New Label With {
                .Dock = DockStyle.Top,
                .Height = 40,
                .Padding = New Padding(8, 6, 8, 0),
                .Text = Tr("Tools.Manage.Hint", "Checked tools are kept on this computer and updated by ""Download / Update"". sqlcmd is downloaded automatically the first time a SQL task needs it.")
            }

            _toolsListView = New ListView With {
                .Dock = DockStyle.Fill,
                .View = View.Details,
                .CheckBoxes = True,
                .FullRowSelect = True,
                .MultiSelect = False,
                .HideSelection = False
            }
            _toolsListView.Columns.Add(Tr("Tools.Manage.ColumnTool", "Tool"), 100)
            _toolsListView.Columns.Add(Tr("Tools.Manage.ColumnDescription", "Description"), 300)
            _toolsListView.Columns.Add(Tr("Tools.Manage.ColumnStatus", "Status"), 90)
            _toolsListView.Columns.Add(Tr("Tools.Manage.ColumnVersion", "Version"), 80)
            _toolsListView.Columns.Add(Tr("Tools.Manage.ColumnSource", "Source"), 150)
            AddHandler _toolsListView.ItemChecked, AddressOf ToolsListView_ItemChecked
            AddHandler _toolsListView.SelectedIndexChanged, Sub(s, e) UpdateButtons()

            _logTextBox = New TextBox With {
                .Dock = DockStyle.Bottom,
                .Height = 110,
                .Multiline = True,
                .ReadOnly = True,
                .ScrollBars = ScrollBars.Vertical,
                .Font = New Font("Consolas", 9)
            }

            Dim buttons As New FlowLayoutPanel With {
                .Dock = DockStyle.Bottom,
                .Height = 40,
                .FlowDirection = FlowDirection.RightToLeft,
                .Padding = New Padding(4)
            }
            _closeButton = New Button With {.Text = Tr("_closeButton.Text", "Close"), .AutoSize = True}
            _removeButton = New Button With {.Text = Tr("Tools.Manage.Remove", "Remove"), .AutoSize = True}
            _installButton = New Button With {.Text = Tr("Tools.Manage.Install", "Download / Update checked tools"), .AutoSize = True}
            AddHandler _closeButton.Click, Sub(s, e) Close()
            AddHandler _removeButton.Click, AddressOf RemoveButton_Click
            AddHandler _installButton.Click, AddressOf InstallButton_Click
            buttons.Controls.AddRange(New Control() {_closeButton, _removeButton, _installButton})

            Text = Tr("Tools.Manage.Title", "LiteTask - Manage Tools")
            Size = New Size(760, 420)
            MinimumSize = New Size(600, 340)
            StartPosition = FormStartPosition.CenterParent
            MaximizeBox = False
            MinimizeBox = False
            AddHandler Me.FormClosing, Sub(s, e) e.Cancel = _busy

            ' Order matters for Dock layout
            Controls.Add(_toolsListView)
            Controls.Add(_hintLabel)
            Controls.Add(_logTextBox)
            Controls.Add(buttons)

            ResumeLayout(False)
        End Sub

        Private Sub LoadTools()
            _loading = True
            Try
                Dim selected = _toolManager.GetSelectedToolIds()
                _toolsListView.Items.Clear()
                For Each tool In _toolManager.Catalog
                    Dim installed = _toolManager.IsInstalled(tool)
                    Dim version = _toolManager.GetInstalledVersion(tool)
                    Dim item As New ListViewItem(tool.DisplayName) With {
                        .Tag = tool,
                        .Checked = selected.Contains(tool.Id)
                    }
                    item.SubItems.Add(T(tool.DescriptionKey, tool.DefaultDescription))
                    item.SubItems.Add(If(installed, Tr("Tools.Manage.Installed", "Installed"), Tr("Tools.Manage.NotInstalled", "Not installed")))
                    item.SubItems.Add(If(installed, If(version = "", "?", version), ""))
                    item.SubItems.Add(tool.Repository)
                    _toolsListView.Items.Add(item)
                Next
            Finally
                _loading = False
            End Try
            UpdateButtons()
        End Sub

        Private Sub UpdateButtons()
            _installButton.Enabled = Not _busy AndAlso _toolsListView.CheckedItems.Count > 0
            Dim current = SelectedTool()
            _removeButton.Enabled = Not _busy AndAlso current IsNot Nothing AndAlso _toolManager.IsInstalled(current)
            _closeButton.Enabled = Not _busy
            _toolsListView.Enabled = Not _busy
        End Sub

        Private Function SelectedTool() As ToolDefinition
            If _toolsListView.SelectedItems.Count = 0 Then Return Nothing
            Return TryCast(_toolsListView.SelectedItems(0).Tag, ToolDefinition)
        End Function

        Private Sub ToolsListView_ItemChecked(sender As Object, e As ItemCheckedEventArgs)
            If _loading Then Return
            Dim ids = _toolsListView.CheckedItems.Cast(Of ListViewItem)().
                Select(Function(i) DirectCast(i.Tag, ToolDefinition).Id).ToList()
            _toolManager.SaveSelectedToolIds(ids)
            UpdateButtons()
        End Sub

        Private Sub AppendLog(message As String)
            If _logTextBox.InvokeRequired Then
                _logTextBox.BeginInvoke(Sub() AppendLog(message))
                Return
            End If
            _logTextBox.AppendText($"{DateTime.Now:HH:mm:ss} {message}{Environment.NewLine}")
        End Sub

        Private Async Sub InstallButton_Click(sender As Object, e As EventArgs)
            _busy = True
            UpdateButtons()
            UseWaitCursor = True
            Try
                Dim tools = _toolsListView.CheckedItems.Cast(Of ListViewItem)().
                    Select(Function(i) DirectCast(i.Tag, ToolDefinition)).ToList()
                Dim failures = 0
                For Each tool In tools
                    Dim result = Await _toolManager.InstallToolAsync(tool, AddressOf AppendLog)
                    If Not result.Success Then failures += 1
                Next
                AppendLog(If(failures = 0,
                             Tr("Tools.Manage.Done", "Done."),
                             String.Format(Tr("Tools.Manage.DoneWithErrors", "Done with {0} error(s); see the log above."), failures)))
            Catch ex As Exception
                AppendLog(ex.Message)
            Finally
                UseWaitCursor = False
                _busy = False
                LoadTools()
            End Try
        End Sub

        Private Sub RemoveButton_Click(sender As Object, e As EventArgs)
            Dim tool = SelectedTool()
            If tool Is Nothing Then Return
            Dim answer = MessageBox.Show(
                String.Format(Tr("Tools.Manage.ConfirmRemove", "Remove {0} from this computer?"), tool.DisplayName),
                Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If answer <> DialogResult.Yes Then Return

            If _toolManager.RemoveTool(tool) Then
                AppendLog(String.Format(Tr("Tools.Manage.Removed", "{0} removed."), tool.DisplayName))
            Else
                AppendLog(String.Format(Tr("Tools.Manage.RemoveFailed", "{0} could not be removed (in use?). See the application log."), tool.DisplayName))
            End If
            LoadTools()
        End Sub
    End Class
End Namespace
