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

#Region " ConsoleWindowWrapper "

' ReSharper disable once CheckNamespace

Namespace DevCase.Core.Application.Console

    ''' <summary>
    ''' Wraps a console window handle and provides a Windows Forms-compatible <see cref="IWin32Window"/> implementation.
    ''' </summary>
    ''' 
    ''' <remarks>
    ''' This class encapsulates an existing window handle (HWND) representing a console window,
    ''' allowing it to be used in contexts that require an <see cref="IWin32Window"/> interface,
    ''' such as setting an owner window for dialogs or message boxes.
    ''' </remarks>
    Public Class ConsoleWindowWrapper : Implements IWin32Window

        ''' <summary>
        ''' Gets the handle (HWND) of the wrapped console window.
        ''' </summary>
        ''' 
        ''' <value>
        ''' The <see cref="IntPtr"/> representing the handle to the console window.
        ''' </value>
        Public ReadOnly Property Handle As IntPtr Implements IWin32Window.Handle

        ''' <summary>
        ''' Initializes a new instance of the <see cref="ConsoleWindowWrapper"/> class
        ''' wrapping the specified console window handle.
        ''' </summary>
        ''' 
        ''' <param name="hwnd">
        ''' The handle to the console window to wrap.
        ''' </param>
        Public Sub New(hwnd As IntPtr)
            Me.Handle = hwnd
        End Sub

    End Class

End Namespace

#End Region
