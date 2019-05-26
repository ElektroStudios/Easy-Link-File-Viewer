' ***********************************************************************
' Author   : ElektroStudios
' Modified : 14-November-2015
' ***********************************************************************

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Shortcut-WindowState "

Namespace DevCase.Core.IO

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Specifies a window state for a shortcut file.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' <see href="https://msdn.microsoft.com/es-es/library/windows/desktop/bb761056%28v=vs.85%29.aspx"/>
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    Public Enum ShortcutWindowState As Integer

        ' *************************************
        ' This enumeration is partially defined
        ' *************************************

        ''' <summary>
        ''' Shortcut Window is at normal state.
        ''' </summary>
        Normal = 1

        ''' <summary>
        ''' Shortcut Window is Maximized.
        ''' </summary>
        Maximized = 3

        ''' <summary>
        ''' Shortcut Window is Minimized.
        ''' </summary>
        Minimized = 7

    End Enum

End Namespace

#End Region
