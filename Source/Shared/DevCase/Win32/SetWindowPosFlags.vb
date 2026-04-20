' ***********************************************************************
' Author   : ElektroStudios
' Modified : 31-March-2017
' ***********************************************************************

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.Diagnostics.CodeAnalysis

#End Region

#Region " SetWindowPos Flags "

' ReSharper disable once CheckNamespace

Namespace DevCase.Win32.Enums

    ''' <summary>
    ''' The window sizing and positioning flags.
    ''' <para></para>
    ''' Flags combination for <c>uFlags</c> parameter of <see cref="NativeMethods.SetWindowPos"/> function.
    ''' </summary>
    '''
    ''' <remarks>
    ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms633545%28v=vs.85%29.aspx"/>
    ''' </remarks>
    <Flags>
    <SuppressMessage("CodeQuality", "IDE0079:Remove unnecessary suppression", Justification:="Required to migrate this code to .NET Core")>
    <SuppressMessage("Design", "CA1069:Enums values should not be duplicated", Justification:="")>
    Public Enum SetWindowPosFlags As UInteger

        ''' <summary>
        ''' Indicates any flag.
        ''' </summary>
        None = &H0UI

        ''' <summary>
        ''' If the calling thread and the thread that owns the window are attached to different input queues, 
        ''' the system posts the request to the thread that owns the window.
        ''' <para></para>
        ''' This prevents the calling thread from blocking its execution while other threads process the request.
        ''' </summary>
        AsyncWindowPos = &H4000

        ''' <summary>
        ''' Prevents generation of the <c>WM_SYNCPAINT</c> message.
        ''' </summary>
        DeferErase = &H2000

        ''' <summary>
        ''' Draws a frame (defined in the window's class description) around the window.
        ''' </summary>
        DrawFrame = &H20

        ''' <summary>
        ''' Applies new frame styles set using the SetWindowLong function.
        ''' <para></para>
        ''' Sends a <c>WM_NCCALCSIZE</c> message to the window, even if the window's size is not being changed.
        ''' <para></para>
        ''' If this flag is not specified, <c>WM_NCCALCSIZE</c> is sent only when the window's size is being changed.
        ''' </summary>
        FrameChanged = &H20

        ''' <summary>
        ''' Hides the window.
        ''' </summary>
        HideWindow = &H80

        ''' <summary>
        ''' Does not activate the window.
        ''' If this flag is not set, the window is activated and moved to the top of
        ''' either the topmost or non-topmost group (depending on the setting of the hWndInsertAfter parameter).
        ''' </summary>
        ''' <remarks>SWP_NOACTIVATE</remarks>
        DontActivate = &H10

        ''' <summary>
        ''' Discards the entire contents of the client area.
        ''' <para></para>
        ''' If this flag is not specified, the valid contents of the client area are saved and copied back 
        ''' into the client area after the window is sized or repositioned.
        ''' </summary>
        DontCopyBits = &H100

        ''' <summary>
        ''' Retains the current position (ignores X and Y parameters).
        ''' </summary>
        IgnoreMove = &H2

        ''' <summary>
        ''' Does not change the owner window's position in the Z order.
        ''' </summary>
        DontChangeOwnerZOrder = &H200

        ''' <summary>
        ''' Does not redraw changes.
        ''' <para></para>
        ''' If this flag is set, no repainting of any kind occurs.
        ''' <para></para>
        ''' This applies to the client area, the nonclient area (including the title bar and scroll bars), 
        ''' and any part of the parent window uncovered as a result of the window being moved.
        ''' <para></para>
        ''' When this flag is set, the application must explicitly invalidate or redraw any parts of 
        ''' the window and parent window that need redrawing.
        ''' </summary>
        DontRedraw = &H8

        ''' <summary>
        ''' Same as the <see cref="SetWindowPosFlags.DontChangeOwnerZOrder"/> flag.
        ''' </summary>
        DontReposition = SetWindowPosFlags.DontChangeOwnerZOrder

        ''' <summary>
        ''' Prevents the window from receiving the <c>WM_WINDOWPOSCHANGING</c> message.
        ''' </summary>
        DontSendChangingEvent = &H400

        ''' <summary>
        ''' Retains the current size (ignores the cx and cy parameters).
        ''' </summary>
        IgnoreResize = &H1

        ''' <summary>
        ''' Retains the current Z order (ignores the <c>hwndInsertAfter</c> parameter).
        ''' </summary>
        IgnoreZOrder = &H4

        ''' <summary>
        ''' Displays the window.
        ''' </summary>
        ShowWindow = &H40

    End Enum

End Namespace

#End Region
