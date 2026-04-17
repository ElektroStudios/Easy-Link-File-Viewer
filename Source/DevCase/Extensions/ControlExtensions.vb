' This source-code is freely distributed as part of "DevCase Class Library .NET Developers".
' 
' Maybe you would like to consider to buy this powerful set of libraries to support me. 
' You can do loads of things with my apis for a big amount of diverse thematics.
' 
' Here is a link to the purchase page:
' https://codecanyon.net/item/elektrokit-class-library-for-net/19260282
' 
' Thank you.

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.ComponentModel
Imports System.ComponentModel.Design
Imports System.Reflection
Imports System.Runtime.CompilerServices

Imports WinForms = System.Windows.Forms

Imports DevCase.Core.Application.UserInterface

#End Region

#Region " Control Extensions "

' ReSharper disable once CheckNamespace

Namespace DevCase.Extensions.ControlExtensions

    ''' <summary>
    ''' Provides extension methods for <see cref="Control"/>.
    ''' </summary>
    <HideModuleName>
    Public Module ControlExtensions


#Region " Public Extension Methods "

        ''' <summary>
        ''' Sets a specified <see cref="ControlStyles"/> flag to
        ''' either <see langword="True"/> or <see langword="False"/> for the source control.
        ''' </summary>
        '''
        ''' <param name="sender">
        ''' The source <see cref="Control"/>.
        ''' </param>
        '''
        ''' <param name="style">
        ''' The <see cref="ControlStyles"/> bit to set.
        ''' </param>
        '''
        ''' <param name="value">
        ''' <see langword="True"/> to apply the specified style to the control; otherwise, <see langword="False"/>.
        ''' </param>
        <DebuggerStepThrough>
        <Extension>
        <EditorBrowsable(EditorBrowsableState.Always)>
        Public Sub SetControlStyle(sender As Control, style As ControlStyles, value As Boolean)

            Dim method As MethodInfo =
                sender.GetType().GetMethod("SetStyle", BindingFlags.NonPublic Or BindingFlags.Instance)

            method.Invoke(sender, {style, value})

        End Sub

        ''' <summary>
        ''' Changes the color appearance of the source <see cref="Control"/> using the specified theme.
        ''' </summary>
        '''
        ''' <param name="ctrl">
        ''' The source <see cref="Control"/>.
        ''' </param>
        '''
        ''' <param name="theme">
        ''' The visual theme.
        ''' </param>
        <DebuggerStepThrough>
        <Extension>
        <EditorBrowsable(EditorBrowsableState.Always)>
        Public Sub SetVisualTheme(ctrl As Control, theme As VisualTheme)

            Select Case theme

                Case VisualTheme.Default
                    ControlExtensions.Internal_SetThemeDefault(ctrl)

                Case VisualTheme.VisualStudioDark
                    ControlExtensions.Internal_SetThemeVisualStudioDark(ctrl)

                Case Else
                    Throw New InvalidEnumArgumentException(NameOf(theme), theme, GetType(VisualTheme))

            End Select

        End Sub

        ''' <summary>
        ''' Iterates through all controls of the specified type within a parent <see cref="Control"/>, 
        ''' optionally recursively, and performs the specified action on each control.
        ''' </summary>
        '''
        ''' <typeparam name="T">
        ''' The type of child controls to iterate through.
        ''' </typeparam>
        ''' 
        ''' <param name="parentControl">
        ''' The parent <see cref="Control"/> whose child controls are to be iterated.
        ''' </param>
        ''' 
        ''' <param name="recursive">
        ''' <see langword="True"/> to iterate recursively through all child controls 
        ''' (i.e., iterate the child controls of child controls); otherwise, <see langword="False"/>.
        ''' </param>
        ''' 
        ''' <param name="action">
        ''' The action to perform on each control.
        ''' </param>
        <DebuggerStepThrough>
        <Extension>
        <EditorBrowsable(EditorBrowsableState.Always)>
        Public Sub ForEachControl(Of T As Control)(parentControl As Control, recursive As Boolean, action As Action(Of T))

            If TypeOf parentControl Is ToolStrip Then
                Throw New InvalidOperationException($"Not allowed. Please use method {NameOf(ToolStripExtensions.ForEachItem)} to iterate items of a {NameOf(ToolStrip)}, {NameOf(StatusStrip)}, {NameOf(MenuStrip)} or {NameOf(Control.ContextMenuStrip)} controls.")
            End If

            If action Is Nothing Then
                Throw New ArgumentNullException(paramName:=NameOf(action), "Action cannot be null.")
            End If

            Dim queue As New Queue(Of Control)

            ' First level items iteration.
            For Each control As Control In parentControl.Controls
                If recursive Then
                    queue.Enqueue(control)
                Else
                    If TypeOf control Is T Then
                        action.Invoke(DirectCast(control, T))
                    End If
                End If
            Next control

            ' Recursive items iteration.
            While queue.Any()
                Dim currentControl As Control = queue.Dequeue()
                If TypeOf currentControl Is T Then
                    action.Invoke(DirectCast(currentControl, T))
                End If

                For Each childControl As Control In currentControl.Controls
                    queue.Enqueue(childControl)
                Next childControl
            End While

            ' PREVIOUS METHODOLOGY. OBSOLETED (IT USES METHOD RECURSION).
            ' -----------------------------------------------------------
            '
            'For Each ctrl As WinForms.Control In parent.Controls.OfType(Of T)
            '    action(ctrl)
            '    If recursive Then
            '        ctrl.ForEachControl(Of T)(recursive, action)
            '    End If
            'Next ctrl

        End Sub


#End Region

#Region " Private Methods "

        ''' <summary>
        ''' Changes the color appearance of the source <see cref="Control"/> to its default appearance.
        ''' </summary>
        '''
        ''' <param name="ctrl">
        ''' The source <see cref="Control"/>.
        ''' </param>
        <DebuggerStepThrough>
        Private Sub Internal_SetThemeDefault(ctrl As Control)

            If ctrl.GetType() = GetType(Button) Then
                With DirectCast(ctrl, Button)
                    .ResetBackColor()
                    .ResetForeColor()
                    .FlatAppearance.BorderColor = Color.Empty
                    .FlatAppearance.BorderSize = 1
                    .UseVisualStyleBackColor = True
                    .UseCompatibleTextRendering = False
                    .FlatStyle = FlatStyle.Standard
                End With

            ElseIf ctrl.GetType() = GetType(ByteViewer) Then
                With DirectCast(ctrl, ByteViewer)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(CheckBox) Then
                With DirectCast(ctrl, CheckBox)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(CheckedListBox) Then
                With DirectCast(ctrl, CheckedListBox)
                    .ResetBackColor()
                    .ResetForeColor()
                    .BorderStyle = BorderStyle.Fixed3D
                End With

            ElseIf ctrl.GetType() = GetType(ComboBox) Then
                With DirectCast(ctrl, ComboBox)
                    .ResetBackColor()
                    .ResetForeColor()
                    .FlatStyle = FlatStyle.Standard
                End With

            ElseIf ctrl.GetType() = GetType(DateTimePicker) Then
                With DirectCast(ctrl, DateTimePicker)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(DataGridView) Then
                With DirectCast(ctrl, DataGridView)
                    .BorderStyle = BorderStyle.FixedSingle
                    .RowTemplate.DefaultCellStyle.BackColor = WinForms.DataGridView.DefaultBackColor
                    .RowTemplate.DefaultCellStyle.ForeColor = WinForms.DataGridView.DefaultForeColor
                    .ResetBackColor()
                    .ResetForeColor()
                    ' .Rows.ForEach(Sub(row As DataGridViewRow) row.DefaultCellStyle.ApplyStyle(.RowTemplate.DefaultCellStyle))
                    '.Rows.ForEach(Sub(row As DataGridViewRow) row.Cells.ForEach(Sub(cell As DataGridViewCell) cell.Style.ApplyStyle(.RowTemplate.DefaultCellStyle)))
                End With

            ElseIf ctrl.GetType() = GetType(FlowLayoutPanel) Then
                With DirectCast(ctrl, FlowLayoutPanel)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(Form) Then
                With DirectCast(ctrl, Form)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(GroupBox) Then
                With DirectCast(ctrl, GroupBox)
                    .ResetBackColor()
                    .ResetForeColor()
                    .FlatStyle = FlatStyle.Standard
                End With

            ElseIf ctrl.GetType() = GetType(HScrollBar) Then
                With DirectCast(ctrl, HScrollBar)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(Label) Then
                With DirectCast(ctrl, Label)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(LinkLabel) Then
                With DirectCast(ctrl, LinkLabel)
                    .ActiveLinkColor = System.Drawing.Color.Red
                    .DisabledLinkColor = System.Drawing.Color.FromArgb(133, 133, 133)
                    .LinkColor = System.Drawing.Color.Blue
                    .VisitedLinkColor = System.Drawing.Color.FromArgb(255, 128, 0, 128)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(ListBox) Then
                With DirectCast(ctrl, ListBox)
                    .ResetBackColor()
                    .ResetForeColor()
                    .BorderStyle = BorderStyle.FixedSingle
                End With

            ElseIf ctrl.GetType() = GetType(ListView) Then
                With DirectCast(ctrl, ListView)
                    .ResetBackColor()
                    .ResetForeColor()
                    .BorderStyle = BorderStyle.Fixed3D
                End With

            ElseIf ctrl.GetType() = GetType(MaskedTextBox) Then
                With DirectCast(ctrl, MaskedTextBox)
                    .ResetBackColor()
                    .ResetForeColor()
                    .BorderStyle = BorderStyle.Fixed3D
                End With

            ElseIf ctrl.GetType() = GetType(MonthCalendar) Then
                With DirectCast(ctrl, MonthCalendar)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(NumericUpDown) Then
                With DirectCast(ctrl, NumericUpDown)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(Panel) Then
                With DirectCast(ctrl, Panel)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(PictureBox) Then
                With DirectCast(ctrl, PictureBox)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(ProgressBar) Then
                With DirectCast(ctrl, ProgressBar)
                    .ResetBackColor()
                    .ResetForeColor()
                    ' RemoveHandler .Paint, AddressOf ControlExtensions.ProgressBar_Paint_VisualStudioDark
                End With
                ControlExtensions.SetControlStyle(ctrl, ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint, False)

            ElseIf ctrl.GetType() = GetType(PropertyGrid) Then
                With DirectCast(ctrl, PropertyGrid)
                    .CategoryForeColor = SystemColors.ControlText
                    .CommandsActiveLinkColor = System.Drawing.Color.Red
                    .CommandsBackColor = SystemColors.Control
                    .CommandsDisabledLinkColor = System.Drawing.Color.FromArgb(133, 133, 133)
                    .CommandsForeColor = SystemColors.ControlText
                    .CommandsLinkColor = System.Drawing.Color.FromArgb(0, 0, 255)
                    .HelpBackColor = SystemColors.Control
                    .HelpForeColor = SystemColors.ControlText
                    .LineColor = SystemColors.InactiveBorder
                    .ViewBackColor = SystemColors.Window
                    .ViewForeColor = SystemColors.WindowText
                    .CategorySplitterColor = SystemColors.Control
                    .CommandsBorderColor = SystemColors.ControlDark
                    .DisabledItemForeColor = SystemColors.GrayText
                    .HelpBorderColor = SystemColors.ControlDark
                    .SelectedItemWithFocusBackColor = SystemColors.Highlight
                    .SelectedItemWithFocusForeColor = SystemColors.HighlightText
                    .ViewBorderColor = SystemColors.ControlDark
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(RadioButton) Then
                With DirectCast(ctrl, RadioButton)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(RichTextBox) Then
                With DirectCast(ctrl, RichTextBox)
                    .ResetBackColor()
                    .ResetForeColor()
                    .BorderStyle = BorderStyle.Fixed3D
                End With

            ElseIf ctrl.GetType() = GetType(SplitContainer) Then
                With DirectCast(ctrl, SplitContainer)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(Splitter) Then
                With DirectCast(ctrl, Splitter)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(SplitterPanel) Then
                With DirectCast(ctrl, SplitterPanel)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(TabControl) Then
                With DirectCast(ctrl, TabControl)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(TabPage) Then
                With DirectCast(ctrl, TabPage)
                    '.ResetBackColor() ' Calling ResetBackColor method does not apply the proper default color.
                    .BackColor = SystemColors.Window
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(TableLayoutPanel) Then
                With DirectCast(ctrl, TableLayoutPanel)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(TextBox) Then
                With DirectCast(ctrl, TextBox)
                    .ResetBackColor()
                    .ResetForeColor()
                    .BorderStyle = BorderStyle.Fixed3D
                End With

            ElseIf ctrl.GetType() = GetType(ToolStrip) Then
                Dim strip As ToolStrip = DirectCast(ctrl, ToolStrip)
                strip.ResetBackColor()
                strip.ResetForeColor()
                ToolStripExtensions.ForEachItem(
                    strip, recursive:=True,
                    Sub(item As ToolStripItem)
                        RemoveHandler item.MouseEnter, AddressOf ControlExtensions.ToolStripItem_MouseEnter_VisualStudioDark
                        RemoveHandler item.MouseLeave, AddressOf ControlExtensions.ToolStripItem_MouseLeave_VisualStudioDark
                        If TypeOf item Is ToolStripMenuItem Then
                            RemoveHandler DirectCast(item, ToolStripMenuItem).DropDownClosed, AddressOf ControlExtensions.ToolStripMenuItem_DropDownClosed_VisualStudioDark
                        End If
                        item.ResetBackColor()
                        item.ResetForeColor()
                    End Sub)
                strip.RenderMode = ToolStripRenderMode.ManagerRenderMode
                strip.Renderer = Nothing

            ElseIf ctrl.GetType() = GetType(MenuStrip) Then
                Dim strip As MenuStrip = DirectCast(ctrl, MenuStrip)
                strip.ResetBackColor()
                strip.ResetForeColor()
                MenuStripExtensions.ForEachItem(
                    strip, recursive:=True,
                    Sub(item As ToolStripItem)
                        RemoveHandler item.MouseEnter, AddressOf ControlExtensions.ToolStripItem_MouseEnter_VisualStudioDark
                        RemoveHandler item.MouseLeave, AddressOf ControlExtensions.ToolStripItem_MouseLeave_VisualStudioDark
                        If TypeOf item Is ToolStripMenuItem Then
                            RemoveHandler DirectCast(item, ToolStripMenuItem).DropDownClosed, AddressOf ControlExtensions.ToolStripMenuItem_DropDownClosed_VisualStudioDark
                        End If
                        item.ResetBackColor()
                        item.ResetForeColor()
                    End Sub)
                strip.RenderMode = ToolStripRenderMode.ManagerRenderMode
                strip.Renderer = Nothing

            ElseIf ctrl.GetType() = GetType(StatusStrip) Then
                Dim strip As StatusStrip = DirectCast(ctrl, StatusStrip)
                strip.ResetBackColor()
                strip.ResetForeColor()
                ToolStripExtensions.ForEachItem(
                    strip, recursive:=True,
                    Sub(item As ToolStripItem)
                        RemoveHandler item.MouseEnter, AddressOf ControlExtensions.ToolStripItem_MouseEnter_VisualStudioDark
                        RemoveHandler item.MouseLeave, AddressOf ControlExtensions.ToolStripItem_MouseLeave_VisualStudioDark
                        If TypeOf item Is ToolStripMenuItem Then
                            RemoveHandler DirectCast(item, ToolStripMenuItem).DropDownClosed, AddressOf ControlExtensions.ToolStripMenuItem_DropDownClosed_VisualStudioDark
                        End If
                        item.ResetBackColor()
                        item.ResetForeColor()
                    End Sub)
                strip.RenderMode = ToolStripRenderMode.ManagerRenderMode
                strip.Renderer = Nothing

            ElseIf ctrl.GetType() = GetType(TrackBar) Then
                With DirectCast(ctrl, TrackBar)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(TreeView) Then
                With DirectCast(ctrl, TreeView)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(VScrollBar) Then
                With DirectCast(ctrl, VScrollBar)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(WebBrowser) Then
                With DirectCast(ctrl, WebBrowser)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            ElseIf ctrl.GetType() = GetType(WebBrowserBase) Then
                With DirectCast(ctrl, WebBrowserBase)
                    .ResetBackColor()
                    .ResetForeColor()
                End With

            Else
                ctrl.ResetBackColor()
                ctrl.ResetForeColor()
                ' Throw New NotImplementedException($"A visual style for the specified control type is not implemented: '{ctrl.GetType().FullName}'")

            End If

        End Sub

        ''' <summary>
        ''' Changes the color appearance of the source <see cref="Control"/> to Visual Studio Dark Theme appearance.
        ''' </summary>
        '''
        ''' <param name="ctrl">
        ''' The source <see cref="Control"/>.
        ''' </param>
        <DebuggerStepThrough>
        Private Sub Internal_SetThemeVisualStudioDark(ctrl As Control)

            If ctrl.GetType() = GetType(Button) Then
                With DirectCast(ctrl, Button)
                    .BackColor = System.Drawing.Color.FromArgb(255, 37, 37, 38)
                    .ForeColor = System.Drawing.Color.Gainsboro
                    .FlatAppearance.BorderColor = Color.DimGray
                    .FlatAppearance.BorderSize = 1
                    .UseVisualStyleBackColor = False
                    .UseCompatibleTextRendering = True
                    .FlatStyle = FlatStyle.Flat
                End With

            ElseIf ctrl.GetType() = GetType(ByteViewer) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(CheckBox) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(CheckedListBox) Then
                With DirectCast(ctrl, CheckedListBox)
                    .BackColor = System.Drawing.Color.FromArgb(255, 37, 37, 38)
                    .ForeColor = System.Drawing.Color.Gainsboro
                    .BorderStyle = BorderStyle.FixedSingle
                End With

            ElseIf ctrl.GetType() = GetType(ComboBox) Then
                With DirectCast(ctrl, ComboBox)
                    .BackColor = System.Drawing.Color.FromArgb(255, 45, 45, 48)
                    .ForeColor = System.Drawing.Color.Gainsboro
                    If .DropDownStyle = ComboBoxStyle.DropDown Then
                        .FlatStyle = FlatStyle.Flat

                        ' Toogling DropDownStyle value forces to recreate the ComboBox handle
                        ' to properly reflect the new BackColor.
                        Dim originalDropDownStyle As ComboBoxStyle = .DropDownStyle
                        .DropDownStyle = ComboBoxStyle.Simple
                        .DropDownStyle = originalDropDownStyle
                    End If
                End With

            ElseIf ctrl.GetType() = GetType(DateTimePicker) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(DataGridView) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro
                With DirectCast(ctrl, DataGridView)
                    .BorderStyle = BorderStyle.FixedSingle
                    .RowTemplate.DefaultCellStyle.BackColor = Color.FromArgb(255, 33, 32, 33)
                    .RowTemplate.DefaultCellStyle.ForeColor = System.Drawing.Color.Gainsboro
                    ' .Rows.ForEach(Sub(row As DataGridViewRow) row.DefaultCellStyle.ApplyStyle(.RowTemplate.DefaultCellStyle))
                    ' .Rows.ForEach(Sub(row As DataGridViewRow) row.Cells.ForEach(Sub(cell As DataGridViewCell) cell.Style.ApplyStyle(.RowTemplate.DefaultCellStyle)))
                End With

            ElseIf ctrl.GetType() = GetType(FlowLayoutPanel) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(Form) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(GroupBox) Then
                With DirectCast(ctrl, GroupBox)
                    .BackColor = System.Drawing.Color.FromArgb(255, 37, 37, 38)
                    .ForeColor = System.Drawing.Color.Gainsboro
                    .FlatStyle = FlatStyle.Flat
                End With

            ElseIf ctrl.GetType() = GetType(HScrollBar) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(Label) Then
                ' ctrl.BackColor = System.Drawing.Color.FromArgb(255, 45, 45, 48)
                ctrl.BackColor = System.Drawing.Color.Transparent
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(LinkLabel) Then
                With DirectCast(ctrl, LinkLabel)
                    .ActiveLinkColor = System.Drawing.Color.IndianRed
                    '  .BackColor = System.Drawing.Color.FromArgb(255, 45, 45, 48)
                    .BackColor = System.Drawing.Color.Transparent
                    .DisabledLinkColor = System.Drawing.Color.FromArgb(133, 133, 133)
                    .ForeColor = System.Drawing.Color.Gainsboro
                    .LinkColor = System.Drawing.Color.FromArgb(255, 0, 122, 204)
                    .VisitedLinkColor = System.Drawing.Color.FromArgb(255, 128, 0, 128)
                End With

            ElseIf ctrl.GetType() = GetType(ListBox) Then
                With DirectCast(ctrl, ListBox)
                    .BackColor = System.Drawing.Color.FromArgb(255, 37, 37, 38)
                    .ForeColor = System.Drawing.Color.Gainsboro
                    .BorderStyle = BorderStyle.FixedSingle
                End With

            ElseIf ctrl.GetType() = GetType(ListView) Then
                With DirectCast(ctrl, ListView)
                    .BackColor = System.Drawing.Color.FromArgb(255, 37, 37, 38)
                    .ForeColor = System.Drawing.Color.Gainsboro
                    .BorderStyle = BorderStyle.FixedSingle
                End With

            ElseIf ctrl.GetType() = GetType(MaskedTextBox) Then
                With DirectCast(ctrl, MaskedTextBox)
                    .BackColor = System.Drawing.Color.FromArgb(255, 37, 37, 38)
                    .ForeColor = System.Drawing.Color.Gainsboro
                    .BorderStyle = BorderStyle.FixedSingle
                End With

            ElseIf ctrl.GetType() = GetType(MonthCalendar) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(ToolStrip) Then
                Dim strip As ToolStrip = DirectCast(ctrl, ToolStrip)
                strip.BackColor = System.Drawing.Color.FromArgb(255, 45, 45, 48)
                strip.ForeColor = System.Drawing.Color.Gainsboro
                ToolStripExtensions.ForEachItem(
                    strip, recursive:=True,
                    Sub(item As ToolStripItem)
                        RemoveHandler item.MouseEnter, AddressOf ControlExtensions.ToolStripItem_MouseEnter_VisualStudioDark
                        RemoveHandler item.MouseLeave, AddressOf ControlExtensions.ToolStripItem_MouseLeave_VisualStudioDark
                        If TypeOf item Is ToolStripMenuItem Then
                            RemoveHandler DirectCast(item, ToolStripMenuItem).DropDownClosed, AddressOf ControlExtensions.ToolStripMenuItem_DropDownClosed_VisualStudioDark
                        End If
                        item.BackColor = System.Drawing.Color.FromArgb(255, 45, 45, 48)
                        item.ForeColor = System.Drawing.Color.Gainsboro
                        AddHandler item.MouseEnter, AddressOf ControlExtensions.ToolStripItem_MouseEnter_VisualStudioDark
                        AddHandler item.MouseLeave, AddressOf ControlExtensions.ToolStripItem_MouseLeave_VisualStudioDark
                        If TypeOf item Is ToolStripMenuItem Then
                            AddHandler DirectCast(item, ToolStripMenuItem).DropDownClosed, AddressOf ControlExtensions.ToolStripMenuItem_DropDownClosed_VisualStudioDark
                        End If
                    End Sub)
                strip.RenderMode = ToolStripRenderMode.System
                strip.Renderer = New ToolStripDarkSystemRenderer()

            ElseIf ctrl.GetType() = GetType(MenuStrip) Then
                Dim strip As MenuStrip = DirectCast(ctrl, MenuStrip)
                strip.BackColor = System.Drawing.Color.FromArgb(255, 45, 45, 48)
                strip.ForeColor = System.Drawing.Color.Gainsboro
                MenuStripExtensions.ForEachItem(
                    strip, recursive:=True,
                    Sub(item As ToolStripItem)
                        RemoveHandler item.MouseEnter, AddressOf ControlExtensions.ToolStripItem_MouseEnter_VisualStudioDark
                        RemoveHandler item.MouseLeave, AddressOf ControlExtensions.ToolStripItem_MouseLeave_VisualStudioDark
                        If TypeOf item Is ToolStripMenuItem Then
                            RemoveHandler DirectCast(item, ToolStripMenuItem).DropDownClosed, AddressOf ControlExtensions.ToolStripMenuItem_DropDownClosed_VisualStudioDark
                        End If
                        item.BackColor = System.Drawing.Color.FromArgb(255, 45, 45, 48)
                        item.ForeColor = System.Drawing.Color.Gainsboro
                        AddHandler item.MouseEnter, AddressOf ControlExtensions.ToolStripItem_MouseEnter_VisualStudioDark
                        AddHandler item.MouseLeave, AddressOf ControlExtensions.ToolStripItem_MouseLeave_VisualStudioDark
                        If TypeOf item Is ToolStripMenuItem Then
                            AddHandler DirectCast(item, ToolStripMenuItem).DropDownClosed, AddressOf ControlExtensions.ToolStripMenuItem_DropDownClosed_VisualStudioDark
                        End If
                    End Sub)
                strip.RenderMode = ToolStripRenderMode.System
                strip.Renderer = New ToolStripDarkSystemRenderer()

            ElseIf ctrl.GetType() = GetType(StatusStrip) Then
                Dim strip As StatusStrip = DirectCast(ctrl, StatusStrip)
                strip.BackColor = System.Drawing.Color.FromArgb(255, 0, 122, 204)
                strip.ForeColor = System.Drawing.Color.Gainsboro
                ToolStripExtensions.ForEachItem(
                    strip, recursive:=True,
                    Sub(item As ToolStripItem)
                        RemoveHandler item.MouseEnter, AddressOf ControlExtensions.ToolStripItem_MouseEnter_VisualStudioDark
                        RemoveHandler item.MouseLeave, AddressOf ControlExtensions.ToolStripItem_MouseLeave_VisualStudioDark
                        If TypeOf item Is ToolStripMenuItem Then
                            RemoveHandler DirectCast(item, ToolStripMenuItem).DropDownClosed, AddressOf ControlExtensions.ToolStripMenuItem_DropDownClosed_VisualStudioDark
                        End If
                        item.BackColor = System.Drawing.Color.FromArgb(255, 0, 122, 204)
                        item.ForeColor = System.Drawing.Color.Gainsboro
                        AddHandler item.MouseEnter, AddressOf ControlExtensions.ToolStripItem_MouseEnter_VisualStudioDark
                        AddHandler item.MouseLeave, AddressOf ControlExtensions.ToolStripItem_MouseLeave_VisualStudioDark
                        If TypeOf item Is ToolStripMenuItem Then
                            AddHandler DirectCast(item, ToolStripMenuItem).DropDownClosed, AddressOf ControlExtensions.ToolStripMenuItem_DropDownClosed_VisualStudioDark
                        End If
                    End Sub)
                strip.RenderMode = ToolStripRenderMode.System
                strip.Renderer = New ToolStripDarkSystemRenderer()

            ElseIf ctrl.GetType() = GetType(NumericUpDown) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(Panel) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(PictureBox) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(ProgressBar) Then
                Dim pb As ProgressBar = DirectCast(ctrl, ProgressBar)
                ctrl.BackColor = Color.FromArgb(255, 30, 30, 30) 'System.Drawing.Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = Color.Black 'System.Drawing.Color.Gainsboro
                ControlExtensions.SetControlStyle(pb, ControlStyles.UserPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint, True)
                ' AddHandler pb.Paint, AddressOf ControlExtensions.ProgressBar_Paint_VisualStudioDark

            ElseIf ctrl.GetType() = GetType(PropertyGrid) Then
                With DirectCast(ctrl, PropertyGrid)
                    .BackColor = System.Drawing.Color.FromArgb(45, 45, 48)
                    .CategoryForeColor = System.Drawing.Color.Silver
                    .CommandsActiveLinkColor = System.Drawing.Color.Red
                    .CommandsBackColor = System.Drawing.Color.FromArgb(45, 45, 48)
                    .CommandsDisabledLinkColor = System.Drawing.Color.FromArgb(133, 133, 133)
                    .CommandsForeColor = System.Drawing.Color.Gainsboro
                    .CommandsLinkColor = System.Drawing.Color.FromArgb(0, 0, 255)
                    .ForeColor = System.Drawing.Color.Gainsboro
                    .HelpBackColor = System.Drawing.Color.FromArgb(45, 45, 48)
                    .HelpForeColor = System.Drawing.Color.Gainsboro
                    .LineColor = System.Drawing.Color.FromArgb(45, 45, 48)
                    .ViewBackColor = System.Drawing.Color.FromArgb(37, 37, 38)
                    .ViewForeColor = System.Drawing.Color.Gainsboro
                    .CategorySplitterColor = System.Drawing.Color.FromArgb(45, 45, 48)
                    .CommandsBorderColor = System.Drawing.Color.Silver
                    .DisabledItemForeColor = System.Drawing.Color.FromArgb(127, 245, 245, 245)
                    .HelpBorderColor = System.Drawing.Color.FromArgb(45, 45, 48)
                    .SelectedItemWithFocusBackColor = SystemColors.Highlight
                    .SelectedItemWithFocusForeColor = SystemColors.HighlightText
                    .ViewBorderColor = System.Drawing.Color.FromArgb(45, 45, 48)
                End With

            ElseIf ctrl.GetType() = GetType(RadioButton) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(RichTextBox) Then
                With DirectCast(ctrl, RichTextBox)
                    .BackColor = System.Drawing.Color.FromArgb(37, 37, 38)
                    .ForeColor = System.Drawing.Color.Gainsboro
                    .BorderStyle = BorderStyle.None
                End With

            ElseIf ctrl.GetType() = GetType(SplitContainer) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(Splitter) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(37, 37, 38)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(SplitterPanel) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(TabControl) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(TabPage) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(TableLayoutPanel) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(TextBox) Then
                With DirectCast(ctrl, TextBox)
                    .BackColor = System.Drawing.Color.FromArgb(255, 37, 37, 38)
                    .ForeColor = System.Drawing.Color.Gainsboro
                    .BorderStyle = BorderStyle.FixedSingle
                End With

            ElseIf ctrl.GetType() = GetType(ToolStrip) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(TrackBar) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(TreeView) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(VScrollBar) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(WebBrowser) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(WebBrowserBase) Then
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro

            Else
                ctrl.BackColor = System.Drawing.Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = System.Drawing.Color.Gainsboro
                ' Throw New NotImplementedException($"A visual style for the specified control type is not implemented: '{ctrl.GetType().FullName}'")

            End If

        End Sub

#End Region

#Region " Event Handlers "

        ''' <summary>
        ''' Handles the <see cref="ToolStripItem.MouseEnter"/> event for <see cref="ToolStripItem"/> controls
        ''' that has the <see cref="VisualTheme.VisualStudioDark"/> style applied.
        ''' </summary>
        '''
        ''' <seealso cref="ControlExtensions.SetVisualTheme"/>
        '''
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        '''
        ''' <param name="e">
        ''' The <see cref="EventArgs"/> instance containing the event data.
        ''' </param>
        Private Sub ToolStripItem_MouseEnter_VisualStudioDark(sender As Object, e As EventArgs)
            Dim item As ToolStripItem = DirectCast(sender, ToolStripItem)
            item.ForeColor = System.Drawing.Color.Black
        End Sub

        ''' <summary>
        ''' Handles the <see cref="ToolStripItem.MouseLeave"/> event for <see cref="ToolStripItem"/> controls
        ''' that has the <see cref="VisualTheme.VisualStudioDark"/> style applied.
        ''' </summary>
        '''
        ''' <seealso cref="ControlExtensions.SetVisualTheme"/>
        '''
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        '''
        ''' <param name="e">
        ''' The <see cref="EventArgs"/> instance containing the event data.
        ''' </param>
        Private Sub ToolStripItem_MouseLeave_VisualStudioDark(sender As Object, e As EventArgs)
            Dim item As ToolStripItem = DirectCast(sender, ToolStripItem)
            item.ForeColor = If(item.Pressed, System.Drawing.Color.Black, System.Drawing.Color.Gainsboro)
        End Sub

        ''' <summary>
        ''' Handles the <see cref="ToolStripMenuItem.DropDownClosed"/> event for <see cref="ToolStripMenuItem"/> controls
        ''' that has the <see cref="VisualTheme.VisualStudioDark"/> style applied.
        ''' </summary>
        '''
        ''' <seealso cref="ControlExtensions.SetVisualTheme"/>
        '''
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        '''
        ''' <param name="e">
        ''' The <see cref="EventArgs"/> instance containing the event data.
        ''' </param>
        Private Sub ToolStripMenuItem_DropDownClosed_VisualStudioDark(sender As Object, e As EventArgs)
            Dim item As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
            item.ForeColor = System.Drawing.Color.Gainsboro
        End Sub

        '''' <summary>
        '''' Handles the <see cref="ProgressBar.Paint"/> event for <see cref="ProgressBar"/> controls
        '''' that has the <see cref="VisualTheme.VisualStudioDark"/> style applied.
        '''' </summary>
        ''''
        '''' <seealso cref="ControlExtensions.SetVisualTheme"/>
        ''''
        '''' <param name="sender">
        '''' The source of the event.
        '''' </param>
        ''''
        '''' <param name="e">
        '''' The <see cref="PaintEventArgs"/> instance containing the event data.
        '''' </param>
        'Private Sub ProgressBar_Paint_VisualStudioDark(sender As Object, e As PaintEventArgs)

        '    Const inset As Integer = 2 ' A single inset value to control the sizing of the inner rect.

        '    Dim pb As ProgressBar = DirectCast(sender, ProgressBar)
        '    Dim width As Integer = pb.Width
        '    Dim height As Integer = pb.Height

        '    Using offscreenImage As Image = UtilImage.CreateSolidcolorBitmap(New Size(width, height), Color.FromArgb(255, 30, 30, 30)),
        '      offscreen As Graphics = Graphics.FromImage(offscreenImage)

        '        Dim rect As New Rectangle(0, 0, width, height)

        '        If ProgressBarRenderer.IsSupported Then
        '            ' ProgressBarRenderer.DrawHorizontalBar(offscreen, rect)
        '        End If

        '        rect.Inflate(New Size(-inset, -inset)) ' Deflate inner rect.
        '        rect.Width = CInt(Math.Truncate(rect.Width * (CDbl(pb.Value) / pb.Maximum)))

        '        If rect.Width <> 0 Then
        '            Dim backColor1 As Color = Color.ForestGreen ' Color.DeepSkyBlue
        '            Dim backColor2 As Color = Color.LightGreen ' Color.LightSkyBlue

        '            Using brush As New LinearGradientBrush(rect, backColor1, backColor2, LinearGradientMode.ForwardDiagonal)
        '                offscreen.FillRectangle(brush, inset, inset, rect.Width, rect.Height)
        '                e.Graphics.DrawImage(offscreenImage, 0, 0)
        '            End Using
        '        End If
        '    End Using

        'End Sub

#End Region

#Region " Internal Classes "

        ''' <summary>
        ''' Handles the painting functionality for <see cref="ToolStrip"/> objects using system colors and a flat visual style.
        ''' </summary>
        ''' 
        ''' <seealso cref="ToolStripSystemRenderer" />
        Friend Class ToolStripDarkSystemRenderer : Inherits ToolStripSystemRenderer

            ''' <summary>
            ''' Initializes a new instance of the <see cref="ToolStripDarkSystemRenderer"/> class.
            ''' </summary>
            Friend Sub New()
            End Sub

            Protected Overrides Sub OnRenderImageMargin(e As ToolStripRenderEventArgs)

                Using brush As New SolidBrush(Color.FromArgb(255, 45, 45, 48))
                    e.Graphics.FillRectangle(brush, e.AffectedBounds)
                End Using
            End Sub

            Protected Overrides Sub OnRenderItemCheck(e As ToolStripItemImageRenderEventArgs)
                Dim checkRect As New Rectangle(
                    e.ImageRectangle.X - 2,
                    e.ImageRectangle.Y - 2,
                    e.ImageRectangle.Width + 4,
                    e.ImageRectangle.Height + 4
                )

                ' Draw check box background.
                Using brush As New SolidBrush(SystemColors.MenuHighlight)
                    e.Graphics.FillRectangle(brush, checkRect)
                End Using

                ' Draw check box border.
                Using pen As New Pen(SystemColors.ControlDarkDark, 1)
                    e.Graphics.DrawRectangle(pen, checkRect)
                End Using

                ' Draw the checkmark or the custom image.
                If e.Image IsNot Nothing Then
                    e.Graphics.DrawImage(e.Image, e.ImageRectangle)
                Else
                    ' Draw a simple checkmark manually.
                    Using pen As New Pen(Color.Gainsboro, 2)
                        Dim x As Integer = e.ImageRectangle.X + 3
                        Dim y As Integer = e.ImageRectangle.Y + (e.ImageRectangle.Height \ 2)
                        e.Graphics.DrawLines(pen, {
                        New Point(x, y),
                        New Point(x + 3, y + 3),
                        New Point(x + 9, y - 4)
                    })
                    End Using
                End If
            End Sub

            Protected Overrides Sub OnRenderSeparator(e As ToolStripSeparatorRenderEventArgs)

                Using pen As New Pen(Color.FromArgb(255, 65, 65, 68), 1)
                    If e.Vertical Then
                        Dim x As Integer = CInt(e.Item.Width \ 2)
                        e.Graphics.DrawLine(pen, New Point(x, 3), New Point(x, e.Item.Height - 4))
                    Else
                        Dim y As Integer = CInt(e.Item.Height \ 2)
                        e.Graphics.DrawLine(pen, New Point(2, y), New Point(e.Item.Width - 4, y))
                    End If
                End Using
            End Sub

            Protected Overrides Sub OnRenderToolStripBackground(e As ToolStripRenderEventArgs)

                Using brush As New SolidBrush(Color.FromArgb(255, 45, 45, 48))
                    e.Graphics.FillRectangle(brush, e.AffectedBounds)
                End Using
            End Sub

            Protected Overrides Sub OnRenderToolStripBorder(e As ToolStripRenderEventArgs)

                Dim y As Integer = e.ToolStrip.Height - 1
                Using pen As New Pen(Color.FromArgb(255, 65, 65, 68), 1)
                    e.Graphics.DrawLine(pen, New Point(0, y), New Point(e.ToolStrip.Width, y))
                End Using
            End Sub

        End Class

#End Region

    End Module

End Namespace

#End Region
