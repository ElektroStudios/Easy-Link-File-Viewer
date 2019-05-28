' This source-code is freely distributed as part of "DevCase for .NET Framework".
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

Imports System.Runtime.InteropServices

#End Region

#Region " CShellLink "

Namespace DevCase.Interop.Unmanaged.Win32.Interfaces

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' <c>CLSID_ShellLink</c> from <c>ShlGuid.h</c> headers.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    <ComImport>
    <ClassInterface(ClassInterfaceType.None)>
    <Guid("00021401-0000-0000-C000-000000000046")>
    Friend Class CShellLink
    End Class

End Namespace

#End Region
