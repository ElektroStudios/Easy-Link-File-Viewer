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

#End Region

#Region " ActiveDesktopApplyMode "

Namespace DevCase.Interop.Unmanaged.Win32.Enums

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Specifies the language format to use for Multilingual User Interface (MUI) functions.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/winnls/nf-winnls-getprocesspreferreduilanguages"/>
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    Friend Enum MuiLanguageMode As Integer

        ''' <summary>
        ''' The input parameter language strings are in language identifier format.
        ''' </summary>
        Id = 4

        ''' <summary>
        ''' The input parameter language strings are in language name format.
        ''' </summary>
        Name = 8

    End Enum

End Namespace

#End Region
