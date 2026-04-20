' ***********************************************************************
' Author   : ElektroStudios
' Modified : 08-October-2024
' ***********************************************************************

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Dialog Box Messages "

' ReSharper disable once CheckNamespace

Namespace DevCase.Win32.Enums

    ''' <summary>
    ''' The system sends or posts a system-defined message when it communicates with an application. 
    ''' <para></para>
    ''' It uses these messages to control the operations of applications and to provide input and other information for applications to process. 
    ''' <para></para>
    ''' An application can also send or post system-defined messages.
    ''' <para></para>
    ''' Applications generally use these messages to control the operation of control windows created by using preregistered window classes.
    ''' </summary>
    '''
    ''' <remarks>
    ''' <see href="https://learn.microsoft.com/en-us/windows/win32/dlgbox/dialog-box-messages"/>
    ''' <para></para>
    ''' The definitions can be found in the Windows SDK file: WinUser.h
    ''' </remarks>
    Public Enum DialogBoxMessages As Integer

        ''' <summary>
        ''' The <see cref="Null"/> message performs no operation.
        ''' <para></para>
        ''' An application sends the <see cref="Null"/> message if it wants to 
        ''' post a message that the recipient window will ignore.
        ''' </summary>
        Null = &H0

        ''' <summary>
        ''' Retrieves the identifier of the default push button control for a dialog box.
        ''' <para></para>
        ''' wParam: This parameter is not used and must be zero.
        ''' <para></para>
        ''' lParam: This parameter is not used and must be zero.
        ''' </summary>
        GetDefId = &H400 + 0

        ''' <summary>
        ''' Changes the identifier of the default push button for a dialog box.
        ''' <para></para>
        ''' wParam: The identifier of a push button control that will become the default.
        ''' <para></para>
        ''' lParam: This parameter is not used and must be zero.
        ''' </summary>
        SetDefId = &H400 + 1

        ''' <summary>
        ''' Repositions a top-level dialog box so that it fits within the desktop area. 
        ''' <para></para>
        ''' An application can send this message to a dialog box after resizing it 
        ''' to ensure that the entire dialog box remains visible. 
        ''' <para></para>
        ''' wParam: This parameter is not used and must be zero.
        ''' <para></para>
        ''' lParam: This parameter is not used and must be zero.
        ''' </summary>
        Reposition = &H400 + 2

    End Enum

End Namespace

#End Region
