Imports System.Runtime.CompilerServices
Imports System.Xml

Namespace LiteTask
    Public Class TranslationManager
        Private Shared _instance As TranslationManager
        Private ReadOnly _translations As Dictionary(Of String, String)
        Private ReadOnly _logger As Logger
        Private ReadOnly _xmlManager As XMLManager
        Private ReadOnly _langPath As String
        Private ReadOnly _currentLanguage As String

        Public Sub New(logger As Logger, xmlManager As XMLManager)
            _translations = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            _logger = logger
            _xmlManager = xmlManager
            _langPath = Path.Combine(Application.StartupPath, "LiteTaskData", "lang")
            _currentLanguage = "fr"
            LoadTranslations()
        End Sub

        Public Function GetTranslation(key As String) As String
            Try
                If _translations.ContainsKey(key) Then
                    Return _translations(key)
                End If
                Return key
            Catch ex As Exception
                _logger?.LogError($"Error getting translation for key '{key}': {ex.Message}")
                Return key
            End Try
        End Function

        Public Function GetTranslation(key As String, defaultValue As String) As String
            Try
                If _translations.ContainsKey(key) Then
                    Return _translations(key)
                End If
                Return defaultValue
            Catch ex As Exception
                _logger?.LogError($"Error getting translation for key '{key}': {ex.Message}")
                Return defaultValue
            End Try
        End Function

        Public Shared Function Initialize(logger As Logger, xmlManager As XMLManager) As TranslationManager
            If _instance Is Nothing Then
                _instance = New TranslationManager(logger, xmlManager)
            End If
            Return _instance
        End Function

        Public Shared ReadOnly Property Instance As TranslationManager
            Get
                If _instance Is Nothing Then
                    Throw New InvalidOperationException("TranslationManager has not been initialized.")
                End If
                Return _instance
            End Get
        End Property

        Private Sub LoadTranslations()
            Try
                If Not Directory.Exists(_langPath) Then
                    Directory.CreateDirectory(_langPath)
                    Return
                End If

                Dim langFile = Path.Combine(_langPath, $"{_currentLanguage}.xml")
                If Not File.Exists(langFile) Then
                    _logger?.LogInfo($"Translation file not found: {langFile}")
                    Return
                End If

                ' Load XML with explicit encoding
                Dim content = File.ReadAllText(langFile, System.Text.Encoding.UTF8)
                Dim doc As New XmlDocument()
                doc.LoadXml(content)

                _translations.Clear()

                ' Load translations
                Dim nodes = doc.SelectNodes("//translation")
                If nodes IsNot Nothing Then
                    For Each node As XmlNode In nodes
                        If node.Attributes("key") IsNot Nothing Then
                            _translations(node.Attributes("key").Value) = node.InnerText.Trim()
                        End If
                    Next
                End If

            Catch ex As Exception
                _logger?.LogError($"Error loading translations: {ex.Message}")
                _translations.Clear()
            End Try
        End Sub

        Public Class ControlTranslator
            Private ReadOnly _translationManager As TranslationManager

            Public Sub New(translationManager As TranslationManager)
                _translationManager = translationManager
            End Sub



            Private Sub TranslateGroupBox(groupBox As GroupBox)
                Try
                    ' Try various translation patterns for GroupBox text
                    Dim translation = TryGetTranslation(
                        $"{groupBox.Name}.Text",                          ' Direct control name
                        $"GroupBox.{groupBox.Name}.Text",                 ' Generic groupbox prefix
                        $"{groupBox.Parent?.Name}.{groupBox.Name}.Text",  ' Parent-based
                        $"GroupBox.{groupBox.Text}",                      ' Generic with current text
                        groupBox.Text                                     ' Current text as fallback
                    )

                    If translation IsNot Nothing Then
                        groupBox.Text = translation
                    End If

                    ' Handle accessibility description if present
                    If Not String.IsNullOrEmpty(groupBox.AccessibleDescription) Then
                        Dim descTranslation = TryGetTranslation(
                            $"{groupBox.Name}.AccessibleDescription",
                            $"GroupBox.{groupBox.Name}.AccessibleDescription",
                            groupBox.AccessibleDescription
                        )
                        If descTranslation IsNot Nothing Then
                            groupBox.AccessibleDescription = descTranslation
                        End If
                    End If

                    ' Handle accessibility name if present
                    If Not String.IsNullOrEmpty(groupBox.AccessibleName) Then
                        Dim nameTranslation = TryGetTranslation(
                            $"{groupBox.Name}.AccessibleName",
                            $"GroupBox.{groupBox.Name}.AccessibleName",
                            groupBox.AccessibleName
                        )
                        If nameTranslation IsNot Nothing Then
                            groupBox.AccessibleName = nameTranslation
                        End If
                    End If

                    ' Special handling for specific GroupBox types based on naming conventions
                    Select Case True
                        Case groupBox.Name.EndsWith("SettingsGroup")
                            HandleSettingsGroupBox(groupBox)
                        Case groupBox.Name.EndsWith("OptionsGroup")
                            HandleOptionsGroupBox(groupBox)
                        Case groupBox.Name.EndsWith("ConfigGroup")
                            HandleConfigGroupBox(groupBox)
                    End Select

                    ' Handle any specific controls within the GroupBox that might need special attention
                    For Each control As Control In groupBox.Controls
                        Select Case True
                            Case TypeOf control Is RadioButton
                                TranslateRadioButton(DirectCast(control, RadioButton), groupBox)
                            Case TypeOf control Is CheckBox
                                TranslateCheckBox(DirectCast(control, CheckBox), groupBox)
                            Case Else
                                TranslateControl(control)
                        End Select
                    Next

                Catch ex As Exception
                    _translationManager._logger?.LogError($"Error translating GroupBox {groupBox.Name}: {ex.Message}")
                End Try
            End Sub

            Private Sub HandleSettingsGroupBox(groupBox As GroupBox)
                ' Special handling for settings groups
                Dim headerTranslation = TryGetTranslation(
                    $"{groupBox.Name}.Header",
                    $"Settings.{groupBox.Name}.Header",
                    $"GroupBox.Settings.Header"
                )
                If headerTranslation IsNot Nothing Then
                    groupBox.Text = headerTranslation
                End If
            End Sub

            Private Sub HandleOptionsGroupBox(groupBox As GroupBox)
                ' Special handling for options groups
                Dim headerTranslation = TryGetTranslation(
                    $"{groupBox.Name}.Header",
                    $"Options.{groupBox.Name}.Header",
                    $"GroupBox.Options.Header"
                )
                If headerTranslation IsNot Nothing Then
                    groupBox.Text = headerTranslation
                End If
            End Sub

            Private Sub HandleConfigGroupBox(groupBox As GroupBox)
                ' Special handling for configuration groups
                Dim headerTranslation = TryGetTranslation(
                    $"{groupBox.Name}.Header",
                    $"Config.{groupBox.Name}.Header",
                    $"GroupBox.Config.Header"
                )
                If headerTranslation IsNot Nothing Then
                    groupBox.Text = headerTranslation
                End If
            End Sub

            Private Sub TranslateRadioButton(radioButton As RadioButton, parentGroup As GroupBox)
                ' Special handling for radio buttons within group boxes
                Dim translation = TryGetTranslation(
                    $"{parentGroup.Name}.{radioButton.Name}.Text",
                    $"{radioButton.Name}.Text",
                    radioButton.Text
                )
                If translation IsNot Nothing Then
                    radioButton.Text = translation
                End If
            End Sub

            Private Sub TranslateCheckBox(checkBox As CheckBox, parentGroup As GroupBox)
                ' Special handling for checkboxes within group boxes
                Dim translation = TryGetTranslation(
                    $"{parentGroup.Name}.{checkBox.Name}.Text",
                    $"{checkBox.Name}.Text",
                    checkBox.Text
                )
                If translation IsNot Nothing Then
                    checkBox.Text = translation
                End If
            End Sub

            Public Sub TranslateControl(control As Control)
                Try
                    If String.IsNullOrEmpty(control.Name) Then Return

                    Select Case True

                        ' Common Controls
                        Case TypeOf control Is Form
                            TranslateForm(DirectCast(control, Form))
                        Case TypeOf control Is UserControl
                            TranslateUserControl(DirectCast(control, UserControl))

                        ' Container Controls
                        Case TypeOf control Is TabControl
                            TranslateTabControl(DirectCast(control, TabControl))
                        Case TypeOf control Is GroupBox
                            TranslateGroupBox(DirectCast(control, GroupBox))
                        Case TypeOf control Is Panel, TypeOf control Is FlowLayoutPanel, TypeOf control Is TableLayoutPanel
                            TranslateContainer(control)

                        ' Menu and Toolbar Controls
                        Case TypeOf control Is MenuStrip
                            TranslateMenuStrip(DirectCast(control, MenuStrip))
                        Case TypeOf control Is ToolStrip
                            TranslateToolStrip(DirectCast(control, ToolStrip))
                        Case TypeOf control Is StatusStrip
                            TranslateStatusStrip(DirectCast(control, StatusStrip))
                        Case TypeOf control Is ContextMenuStrip
                            TranslateContextMenuStrip(DirectCast(control, ContextMenuStrip))

                        ' Data Controls
                        Case TypeOf control Is DataGridView
                            TranslateDataGridView(DirectCast(control, DataGridView))
                        Case TypeOf control Is ListView
                            TranslateListView(DirectCast(control, ListView))

                        ' Input Controls
                        Case TypeOf control Is TextBox, TypeOf control Is RichTextBox
                            TranslateTextBox(control)
                        Case TypeOf control Is ComboBox
                            TranslateComboBox(DirectCast(control, ComboBox))
                        Case TypeOf control Is CheckedListBox
                            TranslateCheckedListBox(DirectCast(control, CheckedListBox))
                        Case TypeOf control Is DateTimePicker
                            TranslateDateTimePicker(DirectCast(control, DateTimePicker))
                        Case TypeOf control Is NumericUpDown
                            TranslateNumericUpDown(DirectCast(control, NumericUpDown))

                        Case Else
                            TranslateGenericControl(control)
                    End Select

                    ' Recursively translate child controls
                    For Each child As Control In control.Controls
                        TranslateControl(child)
                    Next

                Catch ex As Exception
                    _translationManager._logger?.LogError($"Error translating control {control.Name}: {ex.Message}")
                End Try
            End Sub

            Private Function TryGetTranslation(ParamArray keys As String()) As String
                For Each key In keys
                    Dim translation = _translationManager.GetTranslation(key)
                    If translation IsNot Nothing AndAlso translation <> key Then
                        Return translation
                    End If
                Next
                Return Nothing
            End Function

            ' Individual Control Translation Methods
            Private Sub TranslateForm(form As Form)
                Dim translation = TryGetTranslation(
                    $"{form.Name}.Text",
                    $"Form.{form.Name}.Text",
                    form.Text
                )
                If translation IsNot Nothing Then form.Text = translation
            End Sub

            Private Sub TranslateUserControl(userControl As UserControl)
                TranslateGenericControl(userControl)
            End Sub

            Private Sub TranslateTabControl(tabControl As TabControl)
                TranslateGenericControl(tabControl)
                For Each tabPage As TabPage In tabControl.TabPages
                    Dim translation = TryGetTranslation(
                        $"{tabPage.Name}.Text",
                        $"{tabControl.Name}.{tabPage.Name}.Text",
                        tabPage.Text
                    )
                    If translation IsNot Nothing Then tabPage.Text = translation
                Next
            End Sub

            Private Sub TranslateContainer(container As Control)
                TranslateGenericControl(container)
            End Sub

            Private Sub TranslateMenuStrip(menuStrip As MenuStrip)
                TranslateGenericControl(menuStrip)
                For Each item As ToolStripItem In menuStrip.Items
                    TranslateToolStripItem(item)
                Next
            End Sub

            Private Sub TranslateToolStrip(toolStrip As ToolStrip)
                TranslateGenericControl(toolStrip)
                For Each item As ToolStripItem In toolStrip.Items
                    TranslateToolStripItem(item)
                Next
            End Sub

            Private Sub TranslateStatusStrip(statusStrip As StatusStrip)
                TranslateGenericControl(statusStrip)
                For Each item As ToolStripItem In statusStrip.Items
                    TranslateToolStripItem(item)
                Next
            End Sub

            Private Sub TranslateContextMenuStrip(contextMenu As ContextMenuStrip)
                TranslateGenericControl(contextMenu)
                For Each item As ToolStripItem In contextMenu.Items
                    TranslateToolStripItem(item)
                Next
            End Sub

            Private Sub TranslateToolStripItem(item As ToolStripItem)
                Dim translation = TryGetTranslation(
                    $"{item.Name}.Text",
                    $"{item.Owner?.Name}.{item.Name}.Text",
                    item.Text
                )
                If translation IsNot Nothing Then item.Text = translation

                If TypeOf item Is ToolStripMenuItem Then
                    Dim menuItem = DirectCast(item, ToolStripMenuItem)
                    For Each subItem As ToolStripItem In menuItem.DropDownItems
                        TranslateToolStripItem(subItem)
                    Next
                End If
            End Sub

            Private Sub TranslateDataGridView(grid As DataGridView)
                TranslateGenericControl(grid)
                For Each column As DataGridViewColumn In grid.Columns
                    Dim translation = TryGetTranslation(
                        $"{column.Name}.HeaderText",
                        $"{grid.Name}.{column.Name}.HeaderText",
                        column.HeaderText
                    )
                    If translation IsNot Nothing Then column.HeaderText = translation
                Next
            End Sub

            Private Sub TranslateListView(listView As ListView)
                TranslateGenericControl(listView)
                For Each column As ColumnHeader In listView.Columns
                    Dim translation = TryGetTranslation(
                        $"{column.Name}.Text",
                        $"{listView.Name}.{column.Name}.Text",
                        column.Text
                    )
                    If translation IsNot Nothing Then column.Text = translation
                Next
            End Sub

            Private Sub TranslateTextBox(textBox As Control)
                TranslateGenericControl(textBox)
            End Sub

            Private Sub TranslateComboBox(comboBox As ComboBox)
                TranslateGenericControl(comboBox)
                For i As Integer = 0 To comboBox.Items.Count - 1
                    Dim item = comboBox.Items(i).ToString()
                    Dim translation = TryGetTranslation(
                        $"{comboBox.Name}.Items.{item}",
                        item
                    )
                    If translation IsNot Nothing Then comboBox.Items(i) = translation
                Next
            End Sub

            Private Sub TranslateCheckedListBox(checkedListBox As CheckedListBox)
                TranslateGenericControl(checkedListBox)
                For i As Integer = 0 To checkedListBox.Items.Count - 1
                    Dim item = checkedListBox.Items(i).ToString()
                    Dim translation = TryGetTranslation(
                        $"{checkedListBox.Name}.Items.{item}",
                        item
                    )
                    If translation IsNot Nothing Then checkedListBox.Items(i) = translation
                Next
            End Sub

            Private Sub TranslateDateTimePicker(dateTimePicker As DateTimePicker)
                TranslateGenericControl(dateTimePicker)
                If dateTimePicker.CustomFormat <> "" Then
                    Dim translation = TryGetTranslation(
                        $"{dateTimePicker.Name}.CustomFormat",
                        $"DateTimePicker.CustomFormat.{dateTimePicker.Format}"
                    )
                    If translation IsNot Nothing Then dateTimePicker.CustomFormat = translation
                End If
            End Sub

            Private Sub TranslateNumericUpDown(numericUpDown As NumericUpDown)
                TranslateGenericControl(numericUpDown)
            End Sub

            Private Sub TranslateGenericControl(control As Control)
                Dim translation = TryGetTranslation(
                    $"{control.Name}.Text",
                    $"{control.Parent?.Name}.{control.Name}.Text",
                    control.Text
                )
                If translation IsNot Nothing Then control.Text = translation
            End Sub

            ' Specialized Tab Translation Methods
            Public Sub TranslateRunTab(runTab As RunTab)
                Try
                    Dim tabPage = runTab.GetTabPage()
                    If tabPage IsNot Nothing Then
                        ' Translate tab text
                        Dim tabTranslation = TryGetTranslation(
                            "_tabPage.RunTab.Text",
                            "RunTab.Text",
                            tabPage.Text
                        )
                        If tabTranslation IsNot Nothing Then tabPage.Text = tabTranslation

                        ' Translate specific RunTab controls
                        For Each control As Control In tabPage.Controls
                            Select Case control.Name
                                Case "_scriptLabel.Text"
                                    TranslateWithKey(control, "_scriptLabel")
                                Case "_outputLabel.Text"
                                    TranslateWithKey(control, "_outputLabel")
                                Case "_credentialLabel.Text"
                                    TranslateWithKey(control, "_credentialLabel")
                                Case "_executionTypeLabel.Text"
                                    TranslateWithKey(control, "_executionTypeLabel")
                                Case "_targetLabel.Text"
                                    TranslateWithKey(control, "_targetLabel")
                                Case "_requiresElevationCheckBox.Text"
                                    TranslateWithKey(control, "_requiresElevationCheckBox")
                                Case "_autoDetectCheckBox.Text"
                                    TranslateWithKey(control, "_autoDetectCheckBox")
                                Case "_runButton.Text"
                                    TranslateWithKey(control, "_runButton")
                                Case "_executionTypeComboBox"
                                    TranslateComboBoxItems(DirectCast(control, ComboBox))
                                Case Else
                                    TranslateControl(control)
                            End Select
                        Next
                    End If
                Catch ex As Exception
                    _translationManager._logger?.LogError($"Error translating RunTab: {ex.Message}")
                End Try
            End Sub

            Public Sub TranslateSqlTab(sqlTab As SqlTab)
                Try
                    Dim tabPage = sqlTab.GetTabPage()
                    If tabPage IsNot Nothing Then
                        ' Translate tab text
                        Dim tabTranslation = TryGetTranslation(
                            "_tabPage.SQLTab.Text",
                            "SQLTab.Text",
                            tabPage.Text
                        )
                        If tabTranslation IsNot Nothing Then tabPage.Text = tabTranslation

                        ' Translate specific SQLTab controls
                        For Each control As Control In tabPage.Controls
                            Select Case control.Name
                                Case "_queryLabel.Text"
                                    TranslateWithKey(control, "SQLTab._queryLabel")
                                Case "_outputLabel.Text"
                                    TranslateWithKey(control, "SQLTab._outputLabel")
                                Case "_serverLabel.Text"
                                    TranslateWithKey(control, "SQLTab._serverLabel")
                                Case "_databaseLabel.Text"
                                    TranslateWithKey(control, "SQLTab._databaseLabel")
                                Case "_queryTypeLabel.Text"
                                    TranslateWithKey(control, "SQLTab._queryTypeLabel")
                                Case "_objectListLabel.Text"
                                    TranslateWithKey(control, "SQLTab._objectListLabel")
                                Case "_credentialLabel.Text"
                                    TranslateWithKey(control, "SQLTab._credentialLabel")
                                Case "_connectButton.Text"
                                    TranslateWithKey(control, "SQLTab._connectButton")
                                Case "_selectAllButton.Text"
                                    TranslateWithKey(control, "SQLTab._selectAllButton")
                                Case "_returnLastEntryButton.Text"
                                    TranslateWithKey(control, "SQLTab._returnLastEntryButton")
                                Case "_objectListComboBox.Text"
                                    TranslateComboBoxItems(DirectCast(control, ComboBox))
                                Case "_queryTypeComboBox.Text"
                                    TranslateComboBoxItems(DirectCast(control, ComboBox))
                                Case Else
                                    TranslateControl(control)
                            End Select

                            ' Handle DataGridView if present
                            If TypeOf control Is DataGridView Then
                                TranslateDataGridView(DirectCast(control, DataGridView))
                            End If
                        Next
                    End If
                Catch ex As Exception
                    _translationManager._logger?.LogError($"Error translating SQLTab: {ex.Message}")
                End Try
            End Sub

            Private Sub TranslateWithKey(control As Control, translationKey As String)
                Dim translation = TryGetTranslation(translationKey)
                If translation IsNot Nothing Then
                    control.Text = translation
                End If
            End Sub

            Private Sub TranslateComboBoxItems(comboBox As ComboBox)
                For i As Integer = 0 To comboBox.Items.Count - 1
                    Dim item = comboBox.Items(i).ToString()
                    Dim translation = TryGetTranslation(
                        $"{comboBox.Name}.Items.{item}",
                        item
                    )
                    If translation IsNot Nothing Then
                        comboBox.Items(i) = translation
                    End If
                Next
            End Sub
        End Class
    End Class

    ' Extension methods for specialized tab translation
    Public Module SpecializedTabTranslationExtensions
        <Extension()>
        Public Sub TranslateRunTab(runTab As RunTab)
            Try
                If TranslationManager.Instance IsNot Nothing Then
                    Dim translator = New TranslationManager.ControlTranslator(TranslationManager.Instance)
                    translator.TranslateRunTab(runTab)
                End If
            Catch ex As Exception
                ApplicationContainer.GetService(Of Logger)()?.LogWarning($"Translation failed for RunTab: {ex.Message}")
            End Try
        End Sub

        <Extension()>
        Public Sub TranslateSqlTab(sqlTab As SqlTab)
            Try
                If TranslationManager.Instance IsNot Nothing Then
                    Dim translator = New TranslationManager.ControlTranslator(TranslationManager.Instance)
                    translator.TranslateSqlTab(sqlTab)
                End If
            Catch ex As Exception
                ApplicationContainer.GetService(Of Logger)()?.LogWarning($"Translation failed for SQLTab: {ex.Message}")
            End Try
        End Sub
    End Module

    ' Extension method for easy form translation
    Public Module TranslationExtensions
        <Extension()>
        Public Sub Translate(form As Form)
            Try
                If TranslationManager.Instance IsNot Nothing Then
                    Dim translator = New TranslationManager.ControlTranslator(TranslationManager.Instance)
                    translator.TranslateControl(form)
                End If
            Catch ex As Exception
                ApplicationContainer.GetService(Of Logger)()?.LogWarning($"Translation failed for form {form.Name}: {ex.Message}")
            End Try
        End Sub
    End Module
End Namespace