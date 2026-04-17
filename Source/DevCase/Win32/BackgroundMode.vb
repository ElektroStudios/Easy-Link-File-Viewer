' ***********************************************************************
' Author   : ElektroStudios
' Modified : 01-June-2019
' ***********************************************************************

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " BackgroundMode "

' ReSharper disable once CheckNamespace

Namespace DevCase.Win32.Enums

    ''' <summary>
    ''' The background mode used by the <see cref="NativeMethods.GetBkMode"/> 
    ''' and <see cref="NativeMethods.SetBkMode"/> functions.
    ''' </summary>
    '''
    ''' <remarks>
    ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/wingdi/nf-wingdi-setbkmode"/>
    ''' </remarks>
    Public Enum BackgroundMode As Integer

        ''' <summary>
        ''' Indicates that on return, the function has failed.
        ''' </summary>
        [Error] = 0

        ''' <summary>
        ''' Background remains untouched.
        ''' </summary>
        Transparent = 1

        ''' <summary>
        ''' Background is filled with the current background color before the text, hatched brush, or pen is drawn.
        ''' </summary>
        Opaque = 2

    End Enum

End Namespace

#End Region
