' ***********************************************************************
' Author   : ElektroStudios
' Modified : 10-November-2015
' ***********************************************************************

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " WindowLong Values "

' ReSharper disable once CheckNamespace

Namespace DevCase.Win32.Enums

    ''' <summary>
    ''' Specifies the zero-based offset to the value to be get or set on a window.
    ''' </summary>
    '''
    ''' <remarks>
    ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/winuser/nf-winuser-getwindowlongptra"/>
    ''' </remarks>
    Public Enum WindowLongValues As Integer

        ''' <summary>
        ''' Retrieves the user data associated with the window.
        ''' <para></para>
        ''' This data is intended for use by the application that created the window.
        ''' <para></para>
        ''' Its value is initially zero.
        ''' </summary>
        UserData = -21

        ''' <summary>
        ''' Retrieves the extended window styles.
        ''' </summary>
        WindowStyleEx = -20

        ''' <summary>
        ''' Retrieves the window styles.
        ''' </summary>
        WindowStyle = -16

        ''' <summary>
        ''' Retrieves the identifier of the window.
        ''' </summary>
        Id = -12

        ''' <summary>
        ''' Retrieves a handle to the parent window, if any.
        ''' </summary>
        HwndParent = -8

        ''' <summary>
        ''' Retrieves a handle to the application instance.
        ''' </summary>
        HInstance = -6

        ''' <summary>
        ''' Retrieves the address of the window procedure, or a handle representing the address of the window procedure.
        ''' <para></para>
        ''' You must use the <c>CallWindowProc</c> function to call the window procedure.
        ''' </summary>
        WndProc = -4

        ''' <summary>
        ''' ( This value is only available when the <c>hwnd</c> parameter identifies a dialog box. )
        ''' <para></para>
        ''' Retrieves the return value of a message processed in the dialog box procedure.
        ''' </summary>
        DlgMsgresult = 0

        ''' <summary>
        ''' ( This value is only available when the <c>hwnd</c> parameter identifies a dialog box. )
        ''' <para></para>
        ''' Retrieves the address of the dialog box procedure, 
        ''' or a handle representing the address of the dialog box procedure.
        ''' <para></para>
        ''' You must use the <c>CallWindowProc</c> function to call the dialog box procedure.
        ''' </summary>
        DlgProc = 4

        ''' <summary>
        ''' ( This value is only available when the <c>hwnd</c> parameter identifies a dialog box. )
        ''' <para></para>
        ''' Retrieves extra information private to the application, such as handles or pointers.
        ''' </summary>
        DlgUser = 8

    End Enum

End Namespace

#End Region
