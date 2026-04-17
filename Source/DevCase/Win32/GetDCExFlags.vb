' ***********************************************************************
' Author   : ElektroStudios
' Modified : 11-October-2024
' ***********************************************************************

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

#End Region

#Region " GetDCExFlags "

' ReSharper disable once CheckNamespace

Namespace DevCase.Win32.Enums

    ''' <summary>
    ''' Flags for <see cref="NativeMethods.GetDCEx"/> function.
    ''' </summary>
    ''' 
    ''' <remarks>
    ''' <see href="https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getdcex"/>
    ''' </remarks>
    <Flags>
    Public Enum GetDCExFlags

        ''' <summary>
        ''' Returns a DC that corresponds to the window rectangle rather than the client rectangle.
        ''' </summary>
        Window = &H1

        ''' <summary>Returns a DC from the cache, rather than the OWNDC or CLASSDC window. 
        ''' <para></para>
        ''' Essentially overrides CS_OWNDC and CS_CLASSDC.
        ''' </summary>
        Cache = &H2

        ''' <summary>
        ''' This flag is ignored.
        ''' </summary>
        NoResetAttrs = &H4

        ''' <summary>
        ''' Excludes the visible regions of all child windows below the window identified by <c>hWnd</c> parameter.
        ''' </summary>
        ClipChildren = &H8

        ''' <summary>
        ''' Excludes the visible regions of all sibling windows above the window identified by <c>hWnd</c> parameter.
        ''' </summary>
        ClipSiblings = &H10

        ''' <summary>
        ''' Uses the visible region of the parent window. 
        ''' <para></para>
        ''' The parent's <see cref="WindowStyles.ClipChildren"/> and CS_PARENTDC style bits are ignored. 
        ''' <para></para>
        ''' The origin is set to the upper-left corner of the window identified by hWnd.
        ''' </summary>
        ParentClip = &H20

        ''' <summary>
        ''' The clipping region identified by <c>hRgnClip</c> parameter is excluded from the visible region of the returned DC.
        ''' </summary>
        ExcludeRegion = &H40

        ''' <summary>
        ''' The clipping region identified by hrgnClip is intersected with the visible region of the returned DC.
        ''' </summary>
        IntersectRegion = &H80

        ''' <summary>
        ''' Returns a region that excludes the window's update region.
        ''' </summary>
        ExcludeUpdate = &H100

        ''' <summary>
        ''' Returns a region that includes the window's update region.
        ''' </summary>
        IntersectUpdate = &H200

        ''' <summary>
        ''' Allows drawing even if there is a <see cref="NativeMethods.LockWindowUpdate"/> call in 
        ''' effect that would otherwise exclude this window. 
        ''' <para></para>
        ''' Used for drawing during tracking.
        ''' </summary>
        LockWindowUpdate = &H400

        ''' <summary>
        ''' When specified with <see cref="GetDCExFlags.IntersectUpdate"/>, causes the device context to be completely validated.
        ''' <para></para>
        ''' Using this function with both <see cref="GetDCExFlags.IntersectUpdate"/> 
        ''' and <see cref="GetDCExFlags.Validate"/> is identical to using the <see cref="NativeMethods.BeginPaint"/> function.
        ''' </summary>
        Validate = &H200000

    End Enum

End Namespace

#End Region
