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

Imports System.Diagnostics.CodeAnalysis
Imports System.Runtime.InteropServices

#End Region

#Region " ShellFileInfo "

Namespace DevCase.Interop.Unmanaged.Win32.Structures

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains information about a file object.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' <see href="https://msdn.microsoft.com/es-es/library/windows/desktop/bb759792(v=vs.85).aspx"/>
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Friend Structure ShellFileInfo

#Region " Fields "

        ''' <summary>
        ''' A handle to the icon that represents the file. 
        ''' <para></para>
        ''' You are responsible for destroying this handle with <see cref="NativeMethods.DestroyIcon"/> 
        ''' when you no longer need it.
        ''' </summary>
        <SuppressMessage("Microsoft.Security", "CA2111:PointersShouldNotBeVisible", Justification:="Visible for API users")>
        Public IconHandle As IntPtr

        ''' <summary>
        ''' The index of the icon image within the system image list.
        ''' </summary>
        Public IconIndex As Integer

        ''' <summary>
        ''' An array of values that indicates the attributes of the file object. 
        ''' </summary>
        Public Attributes As UInteger

        ''' <summary>
        ''' A string that contains the name of the file as it appears in the Windows Shell, 
        ''' or the path and file name of the file that contains the icon representing the file.
        ''' </summary>
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public DisplayName As String

        ''' <summary>
        ''' A string that describes the type of file.
        ''' </summary>
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)>
        Public TypeName As String

#End Region

    End Structure

End Namespace

#End Region
