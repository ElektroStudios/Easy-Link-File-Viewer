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
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security
Imports System.Text
Imports DevCase.Interop.Unmanaged.Win32.Enums
Imports DevCase.Interop.Unmanaged.Win32.Structures

#End Region

#Region " NativeMethods "

Namespace DevCase.Interop.Unmanaged.Win32

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Platform Invocation methods (P/Invoke), access unmanaged code.
    ''' <para></para>
    ''' This class does not suppress stack walks for unmanaged code permission.
    ''' <see cref="Global.System.Security.SuppressUnmanagedCodeSecurity"/> must not be applied to this class.
    ''' <para></para>
    ''' This class is for methods that can be used anywhere because a stack walk will be performed.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' <see href="https://msdn.microsoft.com/en-us/library/ms182161.aspx"/>
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    Friend NotInheritable Class NativeMethods

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Prevents a default instance of the <see cref="NativeMethods"/> class from being created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerNonUserCode>
        Private Sub New()
        End Sub

#End Region

#Region " Shell32.dll "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves information about an object in the file system, 
        ''' such as a file, folder, directory, or drive root.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/es-es/library/windows/desktop/bb762179(v=vs.85).aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="path">
        ''' A pointer to a null-terminated string of maximum length <c>MAX_PATH</c> that contains the path and file name. 
        ''' Both absolute and relative paths are valid. 
        ''' <para></para>
        ''' If the <paramref name="flags"/> parameter includes the <see cref="SHGetFileInfoFlags.PIDL"/> flag, 
        ''' this parameter must be the address of an <c>ITEMIDLIST</c> (<c>PIDL</c>) structure that contains the 
        ''' list of item identifiers that uniquely identifies the file within the Shell's namespace. 
        ''' The <c>PIDL</c> must be a fully qualified <c>PIDL</c>. Relative <c>PIDLs</c> are not allowed.
        ''' <para></para>
        ''' If the uFlags parameter includes the <see cref="SHGetFileInfoFlags.UseFiles"/> flag, 
        ''' this parameter does not have to be a valid file name. 
        ''' The function will proceed as if the file exists with the specified name and with the 
        ''' file s passed in the <paramref name="fileAttribs"/> parameter. 
        ''' This allows you to obtain information about a file type by passing just the extension for 
        ''' <paramref name="path"/> parameter and passing FILE__NORMAL in <paramref name="fileAttribs"/> parameter
        ''' </param>
        ''' 
        ''' <param name="fileAttribs">
        ''' A combination of one or more file  flags. 
        ''' <para></para>
        ''' If <paramref name="flags"/> parameter does not include the 
        ''' <see cref="SHGetFileInfoFlags.UseFiles"/> flag, this parameter is ignored.
        ''' </param>
        ''' 
        ''' <param name="refShellFileInfo">
        ''' Pointer to a <see cref="ShellFileInfo"/> structure to receive the file information.
        ''' </param>
        ''' 
        ''' <param name="size">
        ''' The size, in bytes, of the <see cref="ShellFileInfo"/> structure pointed to by the 
        ''' <paramref name="refShellFileInfo"/> parameter.
        ''' </param>
        ''' 
        ''' <param name="flags">
        ''' The flags that specify the file information to retrieve.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' Returns a value whose meaning depends on the <paramref name="flags"/> parameter.
        ''' <para></para>
        ''' If <paramref name="flags"/> parameter does not contain <see cref="SHGetFileInfoFlags.ExeType"/> 
        ''' or <see cref="SHGetFileInfoFlags.SysIconIndex"/>, the return value is nonzero if successful, or zero otherwise.
        ''' <para></para>
        ''' If <paramref name="flags"/> parameter contains the <see cref="SHGetFileInfoFlags.ExeType"/> flag, 
        ''' the return value specifies the type of the executable.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <SuppressUnmanagedCodeSecurity>
        <DllImport("Shell32.dll", CharSet:=CharSet.Auto, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function SHGetFileInfo(ByVal path As String,
               <MarshalAs(UnmanagedType.U4)> ByVal fileAttribs As FileAttributes,
                                             ByRef refShellFileInfo As ShellFileInfo,
               <MarshalAs(UnmanagedType.U4)> ByVal size As UInteger,
               <MarshalAs(UnmanagedType.U4)> ByVal flags As SHGetFileInfoFlags
        ) As IntPtr
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves the display name of an item identified by its IDList.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/shobjidl_core/nf-shobjidl_core-shgetnamefromidlist"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidl">
        ''' A PIDL that identifies the item.
        ''' </param>
        ''' 
        ''' <param name="sigdnName">
        ''' A value from the <see cref="ShellItemGetDisplayName"/> enumeration that specifies the type of display name to retrieve.
        ''' </param>
        ''' 
        ''' <param name="name">
        ''' A value that, when this function returns successfully, receives the retrieved display name.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If this function succeeds, it returns <see cref="HResult.S_OK"/>. 
        ''' <para></para>
        ''' Otherwise, it returns an <see cref="HResult"/> error code.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("Shell32.dll", CharSet:=CharSet.Auto, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Public Shared Function SHGetNameFromIDList(ByVal pidl As IntPtr,
                                                   ByVal sigdn As ShellItemGetDisplayName,
                                             <Out> ByRef name As StringBuilder
        ) As HResult
        End Function

#End Region

#Region " User32.dll "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Destroys an icon and frees any memory the icon occupied.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms648063%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hIcon">
        ''' A handle to the icon to be destroyed.
        ''' <para></para>
        ''' The icon must not be in use. 
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the function succeeds, the return value is <see langword="True"/>.
        ''' <para></para>
        ''' If the function fails, the return value is <see langword="False"/>.
        ''' <para></para>
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error()"/>. 
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <SuppressUnmanagedCodeSecurity>
        <DllImport("User32.dll", SetLastError:=True)>
        Friend Shared Function DestroyIcon(ByVal hIcon As IntPtr
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

#End Region

    End Class

End Namespace

#End Region
