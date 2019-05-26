' ***********************************************************************
' Author   : ElektroStudios
' Modified : 10-May-2016
' ***********************************************************************

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
    Public Class CShellLink
    End Class

End Namespace

#End Region
