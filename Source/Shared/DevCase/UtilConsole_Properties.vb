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

Imports DevCase.Win32

#End Region

#Region " CLI Util "

' ReSharper disable once CheckNamespace

Namespace DevCase.Core.Application.Console

    Partial Public NotInheritable Class UtilConsole

#Region " Properties "

        ''' <summary>
        ''' Gets the window handle of the console associated with the calling process.
        ''' </summary>
        '''
        ''' <value>
        ''' The window handle of the console associated with the calling process.
        ''' </value>
        Public Shared ReadOnly Property Handle As IntPtr
            <DebuggerStepThrough>
            Get
                Return NativeMethods.GetConsoleWindow
            End Get
        End Property

#End Region

    End Class

End Namespace

#End Region
