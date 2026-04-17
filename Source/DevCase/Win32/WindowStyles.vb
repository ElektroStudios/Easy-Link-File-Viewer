' ***********************************************************************
' Author   : ElektroStudios
' Modified : 11-October-2024
' ***********************************************************************

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Window Styles "

' ReSharper disable once CheckNamespace

Namespace DevCase.Win32.Enums

    ''' <summary>
    ''' The following styles can be specified wherever a window style is required.
    ''' <para></para>
    ''' After the control has been created, these styles cannot be modified, except as noted.
    ''' </summary>
    '''
    ''' <remarks>
    ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms632600%28v=vs.85%29.aspx"/>
    ''' </remarks>
    <Flags>
    Public Enum WindowStyles As UInteger

        ''' <summary>
        ''' Default window style. Same assee cref="WindowStyles.Overlapped"/>  .
        ''' </summary>
        [Default] = Overlapped ' WS_OVERLAPPED

        ''' <summary>
        ''' The window is an overlapped window. An overlapped window has a title bar and a border.
        ''' </summary>
        Overlapped = &H0

        ''' <summary>
        ''' The window is a pop-up window.
        ''' <para></para>
        ''' This style cannot be used with the <see cref="WindowStyles.Child"/> style.
        ''' </summary>
        Popup = &H80000000UI

        ''' <summary>
        ''' The window is a child window.
        ''' <para></para>
        ''' A window with this style cannot have a menu bar.
        ''' <para></para>
        ''' This style cannot be used with the <see cref="WindowStyles.Popup"/> style.
        ''' </summary>
        Child = &H40000000

        ''' <summary>
        ''' The window is initially minimized.
        ''' </summary>
        Minimize = &H20000000

        ''' <summary>
        ''' The window is initially visible.
        ''' <para></para>
        ''' This style can be turned on and off by using the 
        ''' <see cref="NativeMethods.ShowWindow"/> 
        ''' or <see cref="NativeMethods.SetWindowPos"/> function.
        ''' </summary>
        Visible = &H10000000

        ''' <summary>
        ''' The window is initially disabled.
        ''' <para></para>
        ''' A disabled window cannot receive input from the user.
        ''' <para></para>
        ''' To change this after a window has been created, use the <c>EnableWindow</c> function.
        ''' </summary>
        Disabled = &H8000000

        ''' <summary>
        ''' Clips child windows relative to each other; 
        ''' that is, when a particular child window receives a <see cref="WindowMessages.WM_Paint"/> message, 
        ''' the <see cref="WindowStyles.ClipSiblings"/> style clips all other overlapping child windows out of the 
        ''' region of the child window to be updated.
        ''' <para></para>
        ''' If <see cref="WindowStyles.ClipSiblings"/> is not specified and child windows overlap, 
        ''' it is possible, when drawing within the client area of a child window, 
        ''' to draw within the client area of a neighboring child window.
        ''' </summary>
        ClipSiblings = &H4000000

        ''' <summary>
        ''' Excludes the area occupied by child windows when drawing occurs within the parent window.
        ''' <para></para>
        ''' This style is used when creating the parent window.
        ''' </summary>
        ClipChildren = &H2000000

        ''' <summary>
        ''' The window is initially maximized.
        ''' </summary>
        Maximize = &H1000000

        ''' <summary>
        ''' The window has a title bar.
        ''' <para></para>
        ''' (includes the <see cref="WindowStyles.Border"/> style).
        ''' </summary>
        Caption = Border Or DlgFrame

        ''' <summary>
        ''' The window has a thin-line border.
        ''' </summary>
        Border = &H800000

        ''' <summary>
        ''' The window has a border of a style typically used with dialog boxes.
        ''' <para></para>
        ''' A window with this style cannot have a title bar.
        ''' </summary>
        DlgFrame = &H400000

        ''' <summary>
        ''' The window has a vertical scroll bar.
        ''' </summary>
        VScroll = &H200000

        ''' <summary>
        ''' The window has a horizontal scroll bar.
        ''' </summary>
        HScroll = &H100000

        ''' <summary>
        ''' The window has a window menu on its title bar.
        ''' <para></para>
        ''' The <see cref="WindowStyles.Caption"/> style must also be specified.
        ''' </summary>
        SysMenu = &H80000

        ''' <summary>
        ''' The window has a sizing border.
        ''' </summary>
        ThickFrame = &H40000

        ''' <summary>
        ''' The window is the first control of a group of controls.
        ''' <para></para>
        ''' The group consists of this first control and all controls defined after it, 
        ''' up to the next control with the <see cref="WindowStyles.Group"/> style.
        ''' <para></para>
        ''' The first control in each group usually has the <see cref="WindowStyles.TabStop"/> style 
        ''' so that the user can move from group to group.
        ''' <para></para>
        ''' The user can subsequently change the keyboard focus from one control in the 
        ''' group to the next control in the group by using the direction keys.
        ''' <para></para>
        ''' You can turn this style on and off to change dialog box navigation.
        ''' <para></para>
        ''' To change this style after a window has been created, 
        ''' use the <see cref="NativeMethods.SetWindowLongPtr"/> function.
        ''' </summary>
        Group = &H20000

        ''' <summary>
        ''' The window is a control that can receive the keyboard focus when the user presses the <c>TAB</c> key.
        ''' <para></para>
        ''' Pressing the <c>TAB</c> key changes the keyboard focus to the next control 
        ''' with the <see cref="WindowStyles.TabStop"/> style.
        ''' <para></para> 
        ''' You can turn this style on and off to change dialog box navigation.
        ''' <para></para>
        ''' To change this style after a window has been created, use the <c>SetWindowLong</c> function.
        ''' <para></para>
        ''' For user-created windows and modeless dialogs to work with tab stops, 
        ''' alter the message loop to call the <c>IsDialogMessage</c> function.
        ''' </summary>
        TabStop = &H10000

        ''' <summary>
        ''' The window has a minimize button.
        ''' <para></para>
        ''' Cannot be combined with the <see cref="WindowStylesEx.ContextHelp"/> extended style.
        ''' <para></para>
        ''' The <see cref="WindowStyles.SysMenu"/> style must also be specified.
        ''' </summary>
        MinimizeBox = &H20000

        ''' <summary>
        ''' The window has a maximize button.
        ''' <para></para>
        ''' Cannot be combined with the <see cref="WindowStylesEx.ContextHelp"/> extended style.
        ''' <para></para>
        ''' The <see cref="WindowStyles.SysMenu"/> style must also be specified.
        ''' </summary>
        MaximizeBox = &H10000

        ''' <summary>
        ''' The window is an overlapped window. An overlapped window has a title bar and a border.
        ''' </summary>
        Tiled = WindowStyles.Overlapped

        ''' <summary>
        ''' The window is initially minimized. Same as the <see cref="WindowStyles.Minimize"/> style.
        ''' </summary>
        Iconic = WindowStyles.Minimize

        ''' <summary>
        ''' The window has a sizing border. Same as the <see cref="WindowStyles.ThickFrame"/> style.
        ''' </summary>
        SizeBox = WindowStyles.ThickFrame

        ''' <summary>
        ''' The window is an overlapped window. Same as the <see cref="WindowStyles.OverlappedWindow"/> style.
        ''' </summary>
        TiledWindow = WindowStyles.OverlappedWindow

        ''' <summary>
        ''' The window is an overlapped window.
        ''' </summary>
        OverlappedWindow = WindowStyles.Overlapped Or WindowStyles.Caption Or
                           WindowStyles.SysMenu Or WindowStyles.ThickFrame Or
                           WindowStyles.MinimizeBox Or WindowStyles.MaximizeBox

        ''' <summary>
        ''' The window is a pop-up window.
        ''' <para></para>
        ''' The <see cref="WindowStyles.Caption"/> and <see cref="WindowStyles.PopupWindow"/> styles 
        ''' must be combined to make the window menu visible.
        ''' </summary>
        PopupWindow = WindowStyles.Popup Or WindowStyles.Border Or WindowStyles.SysMenu

        ''' <summary>
        ''' Same as the <see cref="WindowStyles.Child"/> style.
        ''' </summary>
        ChildWindow = WindowStyles.Child

    End Enum

End Namespace

#End Region
