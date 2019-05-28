' This source-code is freely distributed as part of "DevCase for .NET Framework".
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
Imports System.Runtime.CompilerServices

#End Region

#Region " Control Extensions "

Namespace DevCase.Core.Extensions

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains custom extension methods to use with the <see cref="Control"/> type.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    <HideModuleName>
    Public Module ControlExtensions

#Region " Public Extension Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Iterate through all the controls in the source <see cref="Control"/> 
        ''' and performs the specified action on each one.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="parent">
        ''' The source <see cref="Control"/>.
        ''' </param>
        ''' 
        ''' <param name="recursive">
        ''' If <see langword="True"/>, iterates through controls of child controls too.
        ''' </param>
        ''' 
        ''' <param name="action">
        ''' An <see cref="Action(Of Control)"/> to perform on each control.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        <Extension>
        <EditorBrowsable(EditorBrowsableState.Always)>
        Public Sub ForEachControl(ByVal parent As Control, ByVal recursive As Boolean, ByVal action As Action(Of Control))
            ControlExtensions.ForEachControl(Of Control)(parent, recursive, action)
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Iterate through all the controls of the specified type in the source <see cref="Control"/> 
        ''' and performs the specified action on each one.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <typeparam name="T">
        ''' The type of the control.
        ''' </typeparam>
        ''' 
        ''' <param name="parent">
        ''' The source <see cref="Control"/>.
        ''' </param>
        ''' 
        ''' <param name="recursive">
        ''' If <see langword="True"/>, iterates through controls of child controls too.
        ''' </param>
        ''' 
        ''' <param name="action">
        ''' An <see cref="Action(Of Control)"/> to perform on each control.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        <Extension>
        <EditorBrowsable(EditorBrowsableState.Always)>
        Public Sub ForEachControl(Of T As Control)(ByVal parent As Control, ByVal recursive As Boolean, ByVal action As Action(Of Control))
            For Each ctrl As Control In parent.Controls.OfType(Of T)
                action(ctrl)
                If (recursive) Then
                    ControlExtensions.ForEachControl(Of T)(ctrl, recursive, action)
                End If
            Next ctrl
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Changes the color appearance of the source <see cref="Control"/> to its default appearance.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="ctrl">
        ''' The source <see cref="Control"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        <Extension>
        <EditorBrowsable(EditorBrowsableState.Always)>
        Public Sub SetThemeDefault(ByVal ctrl As Control)

            If ctrl.GetType() = GetType(Button) Then
                ctrl.BackColor = Button.DefaultBackColor
                ctrl.ForeColor = Button.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(CheckBox) Then
                ctrl.BackColor = CheckBox.DefaultBackColor
                ctrl.ForeColor = CheckBox.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(CheckedListBox) Then
                ctrl.BackColor = CheckedListBox.DefaultBackColor
                ctrl.ForeColor = CheckedListBox.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(ComboBox) Then
                ctrl.BackColor = ComboBox.DefaultBackColor
                ctrl.ForeColor = ComboBox.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(DateTimePicker) Then
                ctrl.BackColor = DateTimePicker.DefaultBackColor
                ctrl.ForeColor = DateTimePicker.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(DataGridView) Then
                ctrl.BackColor = DataGridView.DefaultBackColor
                ctrl.ForeColor = DataGridView.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(FlowLayoutPanel) Then
                ctrl.BackColor = FlowLayoutPanel.DefaultBackColor
                ctrl.ForeColor = FlowLayoutPanel.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(Form) Then
                ctrl.BackColor = Form.DefaultBackColor
                ctrl.ForeColor = Form.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(GroupBox) Then
                ctrl.BackColor = GroupBox.DefaultBackColor
                ctrl.ForeColor = GroupBox.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(HScrollBar) Then
                ctrl.BackColor = HScrollBar.DefaultBackColor
                ctrl.ForeColor = HScrollBar.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(Label) Then
                ctrl.BackColor = Label.DefaultBackColor
                ctrl.ForeColor = Label.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(LinkLabel) Then
                With DirectCast(ctrl, LinkLabel)
                    .ActiveLinkColor = Color.Red
                    .BackColor = LinkLabel.DefaultBackColor
                    .DisabledLinkColor = Color.FromArgb(133, 133, 133)
                    .ForeColor = LinkLabel.DefaultForeColor
                    .LinkColor = Color.Blue
                    .VisitedLinkColor = Color.FromArgb(255, 128, 0, 128)
                End With

            ElseIf ctrl.GetType() = GetType(ListBox) Then
                ctrl.BackColor = ListBox.DefaultBackColor
                ctrl.ForeColor = ListBox.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(ListView) Then
                ctrl.BackColor = ListView.DefaultBackColor
                ctrl.ForeColor = ListView.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(MaskedTextBox) Then
                ctrl.BackColor = MaskedTextBox.DefaultBackColor
                ctrl.ForeColor = MaskedTextBox.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(MonthCalendar) Then
                ctrl.BackColor = MonthCalendar.DefaultBackColor
                ctrl.ForeColor = MonthCalendar.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(MenuStrip) Then
                ctrl.BackColor = MenuStrip.DefaultBackColor
                ctrl.ForeColor = MenuStrip.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(NumericUpDown) Then
                ctrl.BackColor = NumericUpDown.DefaultBackColor
                ctrl.ForeColor = NumericUpDown.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(panel) Then
                ctrl.BackColor = Panel.DefaultBackColor
                ctrl.ForeColor = Panel.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(PictureBox) Then
                ctrl.BackColor = PictureBox.DefaultBackColor
                ctrl.ForeColor = PictureBox.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(ProgressBar) Then
                ctrl.BackColor = ProgressBar.DefaultBackColor
                ctrl.ForeColor = ProgressBar.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(PropertyGrid) Then
                With DirectCast(ctrl, PropertyGrid)
                    .BackColor = PropertyGrid.DefaultBackColor
                    .CategoryForeColor = SystemColors.ControlText
                    .CategorySplitterColor = SystemColors.Control
                    .CommandsActiveLinkColor = Color.Red
                    .CommandsBackColor = SystemColors.Control
                    .CommandsBorderColor = SystemColors.ControlDark
                    .CommandsDisabledLinkColor = Color.FromArgb(133, 133, 133)
                    .CommandsForeColor = SystemColors.ControlText
                    .CommandsLinkColor = Color.FromArgb(0, 0, 255)
                    .DisabledItemForeColor = SystemColors.GrayText
                    .ForeColor = PropertyGrid.DefaultForeColor
                    .HelpBackColor = SystemColors.Control
                    .HelpBorderColor = SystemColors.ControlDark
                    .HelpForeColor = SystemColors.ControlText
                    .LineColor = SystemColors.InactiveBorder
                    .SelectedItemWithFocusBackColor = SystemColors.Highlight
                    .SelectedItemWithFocusForeColor = SystemColors.HighlightText
                    .ViewBackColor = SystemColors.Window
                    .ViewBorderColor = SystemColors.ControlDark
                    .ViewForeColor = SystemColors.WindowText
                End With

            ElseIf ctrl.GetType() = GetType(RadioButton) Then
                ctrl.BackColor = RadioButton.DefaultBackColor
                ctrl.ForeColor = RadioButton.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(RichTextBox) Then
                ctrl.BackColor = RichTextBox.DefaultBackColor
                ctrl.ForeColor = RichTextBox.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(SplitContainer) Then
                ctrl.BackColor = SplitContainer.DefaultBackColor
                ctrl.ForeColor = SplitContainer.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(Splitter) Then
                ctrl.BackColor = Splitter.DefaultBackColor
                ctrl.ForeColor = Splitter.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(SplitterPanel) Then
                ctrl.BackColor = SplitterPanel.DefaultBackColor
                ctrl.ForeColor = SplitterPanel.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(StatusStrip) Then
                ctrl.BackColor = StatusStrip.DefaultBackColor
                ctrl.ForeColor = StatusStrip.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(TabControl) Then
                ctrl.BackColor = TabControl.DefaultBackColor
                ctrl.ForeColor = TabControl.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(TabPage) Then
                ctrl.BackColor = TabPage.DefaultBackColor
                ctrl.ForeColor = TabPage.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(TableLayoutPanel) Then
                ctrl.BackColor = TableLayoutPanel.DefaultBackColor
                ctrl.ForeColor = TableLayoutPanel.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(TextBox) Then
                ctrl.BackColor = TextBox.DefaultBackColor
                ctrl.ForeColor = TextBox.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(ToolStrip) Then
                ctrl.BackColor = ToolStrip.DefaultBackColor
                ctrl.ForeColor = ToolStrip.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(TrackBar) Then
                ctrl.BackColor = TrackBar.DefaultBackColor
                ctrl.ForeColor = TrackBar.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(TreeView) Then
                ctrl.BackColor = TreeView.DefaultBackColor
                ctrl.ForeColor = TreeView.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(VScrollBar) Then
                ctrl.BackColor = VScrollBar.DefaultBackColor
                ctrl.ForeColor = VScrollBar.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(WebBrowser) Then
                ctrl.BackColor = WebBrowser.DefaultBackColor
                ctrl.ForeColor = WebBrowser.DefaultForeColor

            ElseIf ctrl.GetType() = GetType(WebBrowserBase) Then
                ctrl.BackColor = WebBrowserBase.DefaultBackColor
                ctrl.ForeColor = WebBrowserBase.DefaultForeColor

            Else
                Throw New NotImplementedException($"A color appearance for the specified control type is not implemented: '{ctrl.GetType().FullName}'")

            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Changes the color appearance of the source <see cref="Control"/> to Visual Studio Dark Theme appearance.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="ctrl">
        ''' The source <see cref="Control"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        <Extension>
        <EditorBrowsable(EditorBrowsableState.Always)>
        Public Sub SetThemeVisualStudioDark(ByVal ctrl As Control)

            If ctrl.GetType() = GetType(Button) Then
                ctrl.BackColor = Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(CheckBox) Then
                ctrl.BackColor = Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(CheckedListBox) Then
                ctrl.BackColor = Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(ComboBox) Then
                ctrl.BackColor = Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(DateTimePicker) Then
                ctrl.BackColor = Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(DataGridView) Then
                ctrl.BackColor = Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(FlowLayoutPanel) Then
                ctrl.BackColor = Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(Form) Then
                ctrl.BackColor = Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(GroupBox) Then
                ctrl.BackColor = Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(HScrollBar) Then
                ctrl.BackColor = Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(Label) Then
                ctrl.BackColor = Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(LinkLabel) Then
                With DirectCast(ctrl, LinkLabel)
                    .ActiveLinkColor = Color.IndianRed
                    .BackColor = Color.FromArgb(255, 45, 45, 48)
                    .DisabledLinkColor = Color.FromArgb(133, 133, 133)
                    .ForeColor = Color.Gainsboro
                    .LinkColor = Color.FromArgb(255, 0, 122, 204)
                    .VisitedLinkColor = Color.FromArgb(255, 128, 0, 128)
                End With

            ElseIf ctrl.GetType() = GetType(ListBox) Then
                ctrl.BackColor = Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(ListView) Then
                ctrl.BackColor = Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(MaskedTextBox) Then
                ctrl.BackColor = Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(MonthCalendar) Then
                ctrl.BackColor = Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(MenuStrip) Then
                ctrl.BackColor = Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(NumericUpDown) Then
                ctrl.BackColor = Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(Panel) Then
                ctrl.BackColor = Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(PictureBox) Then
                ctrl.BackColor = Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(ProgressBar) Then
                ctrl.BackColor = Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(PropertyGrid) Then
                With DirectCast(ctrl, PropertyGrid)
                    .BackColor = Color.FromArgb(45, 45, 48)
                    .CategoryForeColor = Color.Silver
                    .CategorySplitterColor = Color.FromArgb(45, 45, 48)
                    .CommandsActiveLinkColor = Color.Red
                    .CommandsBackColor = Color.FromArgb(45, 45, 48)
                    .CommandsBorderColor = Color.Silver
                    .CommandsDisabledLinkColor = Color.FromArgb(133, 133, 133)
                    .CommandsForeColor = Color.Gainsboro
                    .CommandsLinkColor = Color.FromArgb(0, 0, 255)
                    .DisabledItemForeColor = Color.FromArgb(127, 245, 245, 245)
                    .ForeColor = Color.Gainsboro
                    .HelpBackColor = Color.FromArgb(45, 45, 48)
                    .HelpBorderColor = Color.FromArgb(45, 45, 48)
                    .HelpForeColor = Color.Gainsboro
                    .LineColor = Color.FromArgb(45, 45, 48)
                    .SelectedItemWithFocusBackColor = SystemColors.Highlight
                    .SelectedItemWithFocusForeColor = SystemColors.HighlightText
                    .ViewBackColor = Color.FromArgb(37, 37, 38)
                    .ViewBorderColor = Color.FromArgb(45, 45, 48)
                    .ViewForeColor = Color.Gainsboro
                End With

            ElseIf ctrl.GetType() = GetType(RadioButton) Then
                ctrl.BackColor = Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(RichTextBox) Then
                ctrl.BackColor = Color.FromArgb(37, 37, 38)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(SplitContainer) Then
                ctrl.BackColor = Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(Splitter) Then
                ctrl.BackColor = Color.FromArgb(37, 37, 38)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(SplitterPanel) Then
                ctrl.BackColor = Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(StatusStrip) Then
                ctrl.BackColor = Color.FromArgb(255, 0, 122, 204)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(TabControl) Then
                ctrl.BackColor = Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(TabPage) Then
                ctrl.BackColor = Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(TableLayoutPanel) Then
                ctrl.BackColor = Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(TextBox) Then
                ctrl.BackColor = Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(ToolStrip) Then
                ctrl.BackColor = Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(TrackBar) Then
                ctrl.BackColor = Color.FromArgb(255, 45, 45, 48)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(TreeView) Then
                ctrl.BackColor = Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(VScrollBar) Then
                ctrl.BackColor = Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(WebBrowser) Then
                ctrl.BackColor = Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = Color.Gainsboro

            ElseIf ctrl.GetType() = GetType(WebBrowserBase) Then
                ctrl.BackColor = Color.FromArgb(255, 37, 37, 38)
                ctrl.ForeColor = Color.Gainsboro

            Else
                MsgBox(ctrl.GetType().IsNotPublic.ToString)


                Throw New NotImplementedException($"A color appearance for the specified control type is not implemented: '{ctrl.GetType().FullName}'")

            End If

        End Sub

#End Region

    End Module

End Namespace

#End Region
