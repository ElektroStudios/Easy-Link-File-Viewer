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

Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Security
Imports System.Text

Imports DevCase.Interop.Unmanaged.Win32.Enums

#End Region

#Region " IShellItem "

Namespace DevCase.Interop.Unmanaged.Win32.Interfaces

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Exposes a method to return either icons or thumbnails for Shell items. 
    ''' <para></para>
    ''' If no thumbnail or icon is available for the requested item, a per-class icon may be provided from the Shell
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/shobjidl_core/nn-shobjidl_core-ishellitemimagefactory"/>
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    <SuppressUnmanagedCodeSecurity>
    <ComImport>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")>
    Friend Interface IShellItem

        <EditorBrowsable(EditorBrowsableState.Never)>
        Function NotImplemented_BindToHandler() As Object

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the parent of an <see cref="IShellItem"/> object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The address of a pointer to the parent of an <see cref="IShellItem"/> interface.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Function GetParent() As IShellItem

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the display name of the <see cref="IShellItem"/> object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="sigdn">
        ''' One of the <see cref="ShellItemGetDisplayName"/> values that indicates how the name should look.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A value that, when this function returns successfully, 
        ''' receives the address of a pointer to the retrieved display name.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Function GetDisplayName(sigdn As ShellItemGetDisplayName) As StringBuilder

        <EditorBrowsable(EditorBrowsableState.Never)>
        Function NotImplemented_GetAttributes() As UInteger

        <EditorBrowsable(EditorBrowsableState.Never)>
        Function NotImplemented_Compare() As Integer

    End Interface

End Namespace

#End Region
