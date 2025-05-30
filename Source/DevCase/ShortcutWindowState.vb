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
