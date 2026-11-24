Imports System.Drawing
Imports System.Reflection

Namespace LiteTask

    ''' <summary>
    ''' Light or dark display. "auto" follows Windows (Settings > Personalization > Colors, light or
    ''' dark mode of the applications); "light" or "dark" force it (Options, Interface/Theme in
    ''' settings.xml; LITETASK_THEME for tests). Theme.Apply is called by each window after its
    ''' translation: dark colors for the window and its controls, dark title bar and scroll bars,
    ''' accent color of Windows for the selection in menus and tool bars. In light mode the system
    ''' colors are kept; only the menus use the accent color. Same model as LiteDeploy.
    ''' </summary>
    Public NotInheritable Class Theme
        Public Const Automatic As String = "auto"
        Public Const Light As String = "light"
        Public Const Dark As String = "dark"
        Public Const ThemeSetting As String = "Theme"

        Private Shared _dark As Boolean?
        Private Shared _accent As Color?

        ''' <summary>"auto", "light" or "dark" (set at startup from settings.xml).</summary>
        Public Shared Property Preference As String = Automatic

        Private Sub New()
        End Sub

        ''' <summary>True when the windows are drawn dark.</summary>
        Public Shared ReadOnly Property IsDark As Boolean
            Get
                If Not _dark.HasValue Then _dark = ResolveDark(If(Environment.GetEnvironmentVariable("LITETASK_THEME"), Preference), AddressOf WindowsAppsUseDarkMode)
                Return _dark.Value
            End Get
        End Property

        ''' <summary>The mode for a preference: "dark", "light", else what Windows uses. Pure logic, unit-tested.</summary>
        Public Shared Function ResolveDark(preference As String, windowsIsDark As Func(Of Boolean)) As Boolean
            If String.Equals(preference, Dark, StringComparison.OrdinalIgnoreCase) Then Return True
            If String.Equals(preference, Light, StringComparison.OrdinalIgnoreCase) Then Return False
            Return windowsIsDark()
        End Function

        ''' <summary>Interface/Theme of settings.xml ("auto" when absent or unreadable).</summary>
        Public Shared Function ReadPreference(xmlManager As XMLManager) As String
            Try
                If xmlManager Is Nothing Then Return Automatic
                Return xmlManager.ReadValue(TranslationManager.SettingsSection, ThemeSetting, Automatic)
            Catch
                Return Automatic
            End Try
        End Function

        ''' <summary>Forgets the mode and the accent read from Windows (after a change of the preference).</summary>
        Public Shared Sub Reset()
            _dark = Nothing
            _accent = Nothing
        End Sub

        Private Shared Function WindowsAppsUseDarkMode() As Boolean
            Dim value = ReadDword("Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme")
            Return value.HasValue AndAlso value.Value = 0
        End Function

        ''' <summary>Accent color of Windows (DWM), else the selection color of the system.</summary>
        Public Shared ReadOnly Property Accent As Color
            Get
                If Not _accent.HasValue Then
                    Dim abgr = ReadDword("Software\Microsoft\Windows\DWM", "AccentColor")
                    _accent = If(abgr.HasValue,
                                 Color.FromArgb(255, abgr.Value And &HFF, (abgr.Value >> 8) And &HFF, (abgr.Value >> 16) And &HFF),
                                 SystemColors.Highlight)
                End If
                Return _accent.Value
            End Get
        End Property

        ' Dark palette (close to the Windows 11 dark applications)
        Public Shared ReadOnly Property Window As Color
            Get
                Return If(IsDark, Color.FromArgb(32, 32, 32), SystemColors.Control)
            End Get
        End Property

        Public Shared ReadOnly Property Surface As Color
            Get
                Return If(IsDark, Color.FromArgb(43, 43, 43), SystemColors.Window)
            End Get
        End Property

        Public Shared ReadOnly Property Text As Color
            Get
                Return If(IsDark, Color.FromArgb(240, 240, 240), SystemColors.ControlText)
            End Get
        End Property

        Public Shared ReadOnly Property Border As Color
            Get
                Return If(IsDark, Color.FromArgb(70, 70, 70), SystemColors.ControlDark)
            End Get
        End Property

        ''' <summary>Text of an error (status bar): readable in both modes.</summary>
        Public Shared ReadOnly Property ErrorColor As Color
            Get
                Return If(IsDark, Color.FromArgb(255, 110, 100), Color.Red)
            End Get
        End Property

        ''' <summary>Dimmed text (disabled tasks): readable in both modes.</summary>
        Public Shared ReadOnly Property DimmedText As Color
            Get
                Return If(IsDark, Color.FromArgb(150, 150, 150), Color.Gray)
            End Get
        End Property

        ''' <summary>A window: its controls, its title bar, and its tool bars and menus held in fields (context menus).</summary>
        Public Shared Sub Apply(form As Form)
            If form Is Nothing Then Return
            If IsDark Then
                form.BackColor = Window
                form.ForeColor = Text
                If form.IsHandleCreated Then
                    DarkTitleBar(form)
                Else
                    AddHandler form.HandleCreated, Sub(s, e) DarkTitleBar(DirectCast(s, Form))
                End If
                For Each c As Control In form.Controls
                    ApplyTo(c)
                Next
            End If
            ApplyToStripFields(form)
        End Sub

        ''' <summary>A control added after its window was themed (tab page created on demand...).</summary>
        Public Shared Sub Apply(control As Control)
            If control Is Nothing Then Return
            Dim form = TryCast(control, Form)
            If form IsNot Nothing Then
                Apply(form)
                Return
            End If
            ApplyTo(control)
            ApplyToStripFields(control)
        End Sub

        ''' <summary>A menu that is not in the Controls of a window (notification area icon, context menu built in code).</summary>
        Public Shared Sub Apply(strip As ToolStrip)
            If strip IsNot Nothing Then ApplyTo(DirectCast(strip, Control))
        End Sub

        Private Shared Sub ApplyToStripFields(owner As Control)
            For Each field As FieldInfo In owner.GetType().GetFields(BindingFlags.Instance Or BindingFlags.NonPublic Or BindingFlags.Public)
                Dim strip = TryCast(field.GetValue(owner), ToolStrip)
                If strip IsNot Nothing Then ApplyTo(DirectCast(strip, Control))
            Next
        End Sub

        Private Shared Sub ApplyTo(c As Control)
            Dim strip = TryCast(c, ToolStrip)
            If strip IsNot Nothing Then
                strip.Renderer = New ToolStripProfessionalRenderer(New Colors()) With {.RoundedEdges = False}
                If IsDark Then
                    strip.BackColor = Window
                    strip.ForeColor = Text
                    For Each item As ToolStripItem In strip.Items
                        ApplyTo(item)
                    Next
                End If
                Return
            End If
            If Not IsDark Then Return

            Dim list = TryCast(c, ListView)
            If list IsNot Nothing AndAlso list.View = View.Details AndAlso Not list.OwnerDraw Then DarkDetailsView(list)

            Dim tabs = TryCast(c, TabControl)
            If tabs IsNot Nothing AndAlso tabs.DrawMode = TabDrawMode.Normal Then DarkTabs(tabs)

            If TypeOf c Is TextBoxBase OrElse TypeOf c Is ListBox OrElse TypeOf c Is ComboBox OrElse TypeOf c Is ListView OrElse
               TypeOf c Is TreeView OrElse TypeOf c Is NumericUpDown OrElse TypeOf c Is DateTimePicker Then
                ' A box already given its own dark colors (log of the update window) keeps them
                If Not IsAlreadyDark(c.BackColor) Then
                    c.BackColor = Surface
                    c.ForeColor = Text
                End If
                Dim combo = TryCast(c, ComboBox)
                If combo IsNot Nothing Then combo.FlatStyle = FlatStyle.Flat
                If TypeOf c Is ListView OrElse TypeOf c Is TreeView OrElse TypeOf c Is TextBoxBase OrElse TypeOf c Is ListBox Then
                    If c.IsHandleCreated Then
                        DarkScrollBars(c)
                    Else
                        AddHandler c.HandleCreated, Sub(s, e) DarkScrollBars(DirectCast(s, Control))
                    End If
                End If
            ElseIf TypeOf c Is DataGridView Then
                Dim grid = DirectCast(c, DataGridView)
                grid.BackgroundColor = Window
                grid.GridColor = Border
                grid.DefaultCellStyle.BackColor = Surface
                grid.DefaultCellStyle.ForeColor = Text
                grid.DefaultCellStyle.SelectionBackColor = Accent
                grid.DefaultCellStyle.SelectionForeColor = Color.White
                grid.ColumnHeadersDefaultCellStyle.BackColor = Window
                grid.ColumnHeadersDefaultCellStyle.ForeColor = Text
                grid.RowHeadersDefaultCellStyle.BackColor = Window
                grid.RowHeadersDefaultCellStyle.ForeColor = Text
                grid.EnableHeadersVisualStyles = False
            ElseIf TypeOf c Is ButtonBase AndAlso Not TypeOf c Is CheckBox AndAlso Not TypeOf c Is RadioButton Then
                Dim b = DirectCast(c, ButtonBase)
                b.FlatStyle = FlatStyle.Flat
                b.FlatAppearance.BorderColor = Border
                b.BackColor = Color.FromArgb(55, 55, 55)
                b.ForeColor = Text
            Else
                ' Containers and labels: dark background; a color chosen by the code (not a system color) is kept
                If TypeOf c Is GroupBox OrElse TypeOf c Is Panel OrElse TypeOf c Is SplitContainer OrElse TypeOf c Is TabPage OrElse
                   TypeOf c Is UserControl OrElse c.BackColor = SystemColors.Control OrElse c.BackColor = SystemColors.Window OrElse
                   c.BackColor = Color.Transparent Then
                    If c.BackColor <> Color.Transparent OrElse Not TypeOf c Is Label Then c.BackColor = Window
                End If
                c.ForeColor = Text
            End If

            For Each child As Control In c.Controls
                ApplyTo(child)
            Next
        End Sub

        Private Shared Function IsAlreadyDark(color As Color) As Boolean
            Return Not color.IsSystemColor AndAlso color.A = 255 AndAlso color.GetBrightness() < 0.25F
        End Function

        Private Shared Sub ApplyTo(item As ToolStripItem)
            item.ForeColor = Text
            Dim drop = TryCast(item, ToolStripDropDownItem)
            If drop IsNot Nothing Then
                drop.DropDown.BackColor = Window
                drop.DropDown.ForeColor = Text
                For Each [sub] As ToolStripItem In drop.DropDownItems
                    ApplyTo([sub])
                Next
            End If
        End Sub

        ''' <summary>
        ''' Background of a row in dark mode: a row colored by the code (state of a task) keeps a dark
        ''' tint of its color, so the text stays light and readable. Pure logic, unit-tested.
        ''' </summary>
        Public Shared Function DarkRowBack(itemBack As Color, surfaceColor As Color) As Color
            If itemBack.IsEmpty OrElse itemBack.A = 0 OrElse itemBack.IsSystemColor OrElse itemBack.ToArgb() = Color.White.ToArgb() Then Return surfaceColor
            ' 30 % of the color over the dark surface
            Return Color.FromArgb(255,
                                  CInt(surfaceColor.R + (CInt(itemBack.R) - surfaceColor.R) * 0.3),
                                  CInt(surfaceColor.G + (CInt(itemBack.G) - surfaceColor.G) * 0.3),
                                  CInt(surfaceColor.B + (CInt(itemBack.B) - surfaceColor.B) * 0.3))
        End Function

        ' A details list drawn by Windows keeps white grid lines and light headers in dark mode:
        ' it is drawn here (headers, rows, selection, check boxes, images, grid lines)
        Private Shared Sub DarkDetailsView(list As ListView)
            Dim gridLines = list.GridLines
            list.GridLines = False
            list.OwnerDraw = True

            AddHandler list.DrawColumnHeader,
                Sub(s, e)
                    Using back As New SolidBrush(Window)
                        e.Graphics.FillRectangle(back, e.Bounds)
                    End Using
                    Using line As New Pen(Border)
                        e.Graphics.DrawLine(line, e.Bounds.Right - 1, e.Bounds.Top + 3, e.Bounds.Right - 1, e.Bounds.Bottom - 4)
                        e.Graphics.DrawLine(line, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1)
                    End Using
                    Dim textBounds As New Rectangle(e.Bounds.X + 6, e.Bounds.Y, e.Bounds.Width - 8, e.Bounds.Height)
                    TextRenderer.DrawText(e.Graphics, e.Header.Text, list.Font, textBounds, Text,
                        TextFormatFlags.SingleLine Or TextFormatFlags.VerticalCenter Or TextFormatFlags.Left Or TextFormatFlags.EndEllipsis Or TextFormatFlags.NoPrefix)
                End Sub

            AddHandler list.DrawItem, Sub(s, e)
                                          ' everything is drawn per sub-item
                                      End Sub

            AddHandler list.DrawSubItem,
                Sub(s, e)
                    Dim item = e.Item
                    Dim bounds = e.Bounds
                    Using brush As New SolidBrush(DarkRowBack(item.BackColor, Surface))
                        e.Graphics.FillRectangle(brush, bounds)
                    End Using
                    If item.Selected Then
                        Using brush As New SolidBrush(Color.FromArgb(110, Accent))
                            e.Graphics.FillRectangle(brush, bounds)
                        End Using
                    End If

                    Dim x = bounds.X + 4
                    If e.ColumnIndex = 0 Then
                        If list.CheckBoxes Then
                            Dim box As New Point(x, bounds.Y + (bounds.Height - 13) \ 2)
                            CheckBoxRenderer.DrawCheckBox(e.Graphics, box, If(item.Checked, CheckBoxState.CheckedNormal, CheckBoxState.UncheckedNormal))
                            x += 17
                        End If
                        Dim images = list.StateImageList
                        Dim index = item.StateImageIndex
                        If (images Is Nothing OrElse index < 0) AndAlso list.SmallImageList IsNot Nothing AndAlso item.ImageIndex >= 0 Then
                            images = list.SmallImageList
                            index = item.ImageIndex
                        End If
                        If images IsNot Nothing AndAlso index >= 0 AndAlso index < images.Images.Count Then
                            e.Graphics.DrawImage(images.Images(index), x, bounds.Y + (bounds.Height - images.ImageSize.Height) \ 2,
                                                 images.ImageSize.Width, images.ImageSize.Height)
                            x += images.ImageSize.Width + 3
                        End If
                    End If

                    Dim fore = e.SubItem.ForeColor
                    If e.ColumnIndex = 0 OrElse item.UseItemStyleForSubItems Then fore = item.ForeColor
                    If fore = SystemColors.WindowText OrElse fore = SystemColors.ControlText OrElse fore.IsEmpty OrElse fore.GetBrightness() < 0.35F Then
                        fore = If(fore = Color.Gray, DimmedText, Text)
                    End If
                    Dim textBounds As New Rectangle(x, bounds.Y, Math.Max(0, bounds.Right - x - 2), bounds.Height)
                    ' SingleLine: a text with a line break stays on the row
                    TextRenderer.DrawText(e.Graphics, e.SubItem.Text, If(item.Font, list.Font), textBounds, fore,
                        TextFormatFlags.SingleLine Or TextFormatFlags.VerticalCenter Or TextFormatFlags.Left Or TextFormatFlags.EndEllipsis Or TextFormatFlags.NoPrefix)

                    If gridLines Then
                        Using line As New Pen(Border)
                            e.Graphics.DrawLine(line, bounds.Left, bounds.Bottom - 1, bounds.Right, bounds.Bottom - 1)
                            e.Graphics.DrawLine(line, bounds.Right - 1, bounds.Top, bounds.Right - 1, bounds.Bottom)
                        End Using
                    End If
                End Sub

            ' Windows still handles the click on the check box: only the row is redrawn
            If list.CheckBoxes Then AddHandler list.ItemChecked, Sub(s, e) list.Invalidate(e.Item.Bounds)
        End Sub

        ' Tabs drawn by Windows stay light in dark mode
        Private Shared Sub DarkTabs(tabs As TabControl)
            tabs.DrawMode = TabDrawMode.OwnerDrawFixed
            AddHandler tabs.DrawItem,
                Sub(s, e)
                    Dim selected = e.Index = tabs.SelectedIndex
                    Using back As New SolidBrush(If(selected, Surface, Window))
                        e.Graphics.FillRectangle(back, e.Bounds)
                    End Using
                    If selected Then
                        Using accentPen As New Pen(Accent, 2)
                            e.Graphics.DrawLine(accentPen, e.Bounds.Left + 2, e.Bounds.Bottom - 1, e.Bounds.Right - 2, e.Bounds.Bottom - 1)
                        End Using
                    End If
                    TextRenderer.DrawText(e.Graphics, tabs.TabPages(e.Index).Text, tabs.Font, e.Bounds, Text,
                        TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPrefix)
                End Sub
        End Sub

        ' Selection in menus and tool bars: the accent color (both modes); dark menus in dark mode
        Private NotInheritable Class Colors
            Inherits ProfessionalColorTable

            Private ReadOnly Property Selected As Color
                Get
                    Return Color.FromArgb(If(IsDark, 110, 60), Accent)
                End Get
            End Property

            Public Overrides ReadOnly Property MenuItemSelected As Color
                Get
                    Return Selected
                End Get
            End Property
            Public Overrides ReadOnly Property MenuItemSelectedGradientBegin As Color
                Get
                    Return Selected
                End Get
            End Property
            Public Overrides ReadOnly Property MenuItemSelectedGradientEnd As Color
                Get
                    Return Selected
                End Get
            End Property
            Public Overrides ReadOnly Property MenuItemPressedGradientBegin As Color
                Get
                    Return If(IsDark, Theme.Surface, MyBase.MenuItemPressedGradientBegin)
                End Get
            End Property
            Public Overrides ReadOnly Property MenuItemPressedGradientEnd As Color
                Get
                    Return If(IsDark, Theme.Surface, MyBase.MenuItemPressedGradientEnd)
                End Get
            End Property
            Public Overrides ReadOnly Property MenuItemBorder As Color
                Get
                    Return Accent
                End Get
            End Property
            Public Overrides ReadOnly Property ButtonSelectedHighlight As Color
                Get
                    Return Selected
                End Get
            End Property
            Public Overrides ReadOnly Property ButtonSelectedGradientBegin As Color
                Get
                    Return Selected
                End Get
            End Property
            Public Overrides ReadOnly Property ButtonSelectedGradientEnd As Color
                Get
                    Return Selected
                End Get
            End Property
            Public Overrides ReadOnly Property ButtonPressedGradientBegin As Color
                Get
                    Return Accent
                End Get
            End Property
            Public Overrides ReadOnly Property ButtonPressedGradientEnd As Color
                Get
                    Return Accent
                End Get
            End Property
            Public Overrides ReadOnly Property ButtonSelectedBorder As Color
                Get
                    Return Accent
                End Get
            End Property
            Public Overrides ReadOnly Property ButtonCheckedHighlight As Color
                Get
                    Return Selected
                End Get
            End Property
            Public Overrides ReadOnly Property ButtonCheckedGradientBegin As Color
                Get
                    Return Selected
                End Get
            End Property
            Public Overrides ReadOnly Property ButtonCheckedGradientMiddle As Color
                Get
                    Return Selected
                End Get
            End Property
            Public Overrides ReadOnly Property ButtonCheckedGradientEnd As Color
                Get
                    Return Selected
                End Get
            End Property
            Public Overrides ReadOnly Property ToolStripDropDownBackground As Color
                Get
                    Return If(IsDark, Theme.Surface, MyBase.ToolStripDropDownBackground)
                End Get
            End Property
            Public Overrides ReadOnly Property ImageMarginGradientBegin As Color
                Get
                    Return If(IsDark, Theme.Surface, MyBase.ImageMarginGradientBegin)
                End Get
            End Property
            Public Overrides ReadOnly Property ImageMarginGradientMiddle As Color
                Get
                    Return If(IsDark, Theme.Surface, MyBase.ImageMarginGradientMiddle)
                End Get
            End Property
            Public Overrides ReadOnly Property ImageMarginGradientEnd As Color
                Get
                    Return If(IsDark, Theme.Surface, MyBase.ImageMarginGradientEnd)
                End Get
            End Property
            Public Overrides ReadOnly Property MenuBorder As Color
                Get
                    Return If(IsDark, Theme.Border, MyBase.MenuBorder)
                End Get
            End Property
            Public Overrides ReadOnly Property MenuStripGradientBegin As Color
                Get
                    Return If(IsDark, Theme.Window, MyBase.MenuStripGradientBegin)
                End Get
            End Property
            Public Overrides ReadOnly Property MenuStripGradientEnd As Color
                Get
                    Return If(IsDark, Theme.Window, MyBase.MenuStripGradientEnd)
                End Get
            End Property
            Public Overrides ReadOnly Property ToolStripGradientBegin As Color
                Get
                    Return If(IsDark, Theme.Window, MyBase.ToolStripGradientBegin)
                End Get
            End Property
            Public Overrides ReadOnly Property ToolStripGradientMiddle As Color
                Get
                    Return If(IsDark, Theme.Window, MyBase.ToolStripGradientMiddle)
                End Get
            End Property
            Public Overrides ReadOnly Property ToolStripGradientEnd As Color
                Get
                    Return If(IsDark, Theme.Window, MyBase.ToolStripGradientEnd)
                End Get
            End Property
            Public Overrides ReadOnly Property ToolStripBorder As Color
                Get
                    Return If(IsDark, Theme.Window, MyBase.ToolStripBorder)
                End Get
            End Property
            Public Overrides ReadOnly Property StatusStripGradientBegin As Color
                Get
                    Return If(IsDark, Theme.Window, MyBase.StatusStripGradientBegin)
                End Get
            End Property
            Public Overrides ReadOnly Property StatusStripGradientEnd As Color
                Get
                    Return If(IsDark, Theme.Window, MyBase.StatusStripGradientEnd)
                End Get
            End Property
            Public Overrides ReadOnly Property SeparatorDark As Color
                Get
                    Return If(IsDark, Theme.Border, MyBase.SeparatorDark)
                End Get
            End Property
            Public Overrides ReadOnly Property SeparatorLight As Color
                Get
                    Return If(IsDark, Theme.Window, MyBase.SeparatorLight)
                End Get
            End Property
        End Class

        ' ---- Win32 ------------------------------------------------------------------

        <DllImport("dwmapi.dll")>
        Private Shared Function DwmSetWindowAttribute(hwnd As IntPtr, attribute As Integer, ByRef value As Integer, size As Integer) As Integer
        End Function

        <DllImport("uxtheme.dll", CharSet:=CharSet.Unicode)>
        Private Shared Function SetWindowTheme(hwnd As IntPtr, appName As String, idList As String) As Integer
        End Function

        <DllImport("user32.dll")>
        Private Shared Function SendMessage(hwnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
        End Function

        Private Const DWMWA_USE_IMMERSIVE_DARK_MODE As Integer = 20
        Private Const DWMWA_USE_IMMERSIVE_DARK_MODE_OLD As Integer = 19   ' Windows 10 before 20H1
        Private Const LVM_GETHEADER As Integer = &H101F

        Private Shared Sub DarkTitleBar(form As Form)
            Try
                Dim isOn = 1
                If DwmSetWindowAttribute(form.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, isOn, 4) <> 0 Then
                    DwmSetWindowAttribute(form.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE_OLD, isOn, 4)
                End If
            Catch
                ' older Windows: light title bar
            End Try
        End Sub

        Private Shared Sub DarkScrollBars(c As Control)
            Try
                SetWindowTheme(c.Handle, "DarkMode_Explorer", Nothing)
                If TypeOf c Is ListView Then
                    Dim header = SendMessage(c.Handle, LVM_GETHEADER, IntPtr.Zero, IntPtr.Zero)
                    If header <> IntPtr.Zero Then SetWindowTheme(header, "DarkMode_ItemsView", Nothing)
                End If
            Catch
            End Try
        End Sub

        Private Shared Function ReadDword(key As String, name As String) As Integer?
            Try
                Using k = Registry.CurrentUser.OpenSubKey(key)
                    Dim v = k?.GetValue(name)
                    If TypeOf v Is Integer Then Return DirectCast(v, Integer)
                End Using
            Catch
            End Try
            Return Nothing
        End Function
    End Class
End Namespace
