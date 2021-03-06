﻿' This source-code is freely distributed as part of "DevCase for .NET Framework".
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

Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security
Imports System.Text

Imports DevCase.Interop.Unmanaged.Win32.Enums
Imports DevCase.Interop.Unmanaged.Win32.Interfaces
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

#Region " Kernel32.dll "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sets the process preferred UI languages for the current process.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/winnls/nf-winnls-setprocesspreferreduilanguages"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="flags">
        ''' Flags identifying the language format to use for the process preferred UI languages.
        ''' </param>
        ''' 
        ''' <param name="languagesBuffer">
        ''' Pointer to a double null-terminated multi-string buffer that contains an ordered, 
        ''' null-delimited list in decreasing order of preference. 
        ''' <para></para>
        ''' If there are more than five languages in the buffer, the function only sets the first five valid languages.
        ''' <para></para>
        ''' Alternatively, this parameter can contain NULL if no language list is required. 
        ''' In this case, the function clears the preferred UI languages for the process.
        ''' </param>
        ''' 
        ''' <param name="refNumLanguages">
        ''' Pointer to the number of languages that has been set in the process language list from the input buffer, 
        ''' up to a maximum of five.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the function succeeds, the return value is <see langword="True"/>. 
        ''' If the function fails, the return value is <see langword="False"/>.
        ''' <para></para>
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error()"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("Kernel32.dll", SetLastError:=True, ExactSpelling:=True, CharSet:=CharSet.Unicode)>
        Friend Shared Function SetProcessPreferredUILanguages(flags As MuiLanguageMode,
                                                              languagesBuffer As String,
                                                        <Out> ByRef refNumLanguages As UInteger
        ) As Boolean
        End Function

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
        Friend Shared Function SHGetFileInfo(path As String,
               <MarshalAs(UnmanagedType.U4)> fileAttribs As FileAttributes,
                                             ByRef refShellFileInfo As ShellFileInfo,
               <MarshalAs(UnmanagedType.U4)> size As UInteger,
               <MarshalAs(UnmanagedType.U4)> flags As SHGetFileInfoFlags
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
        ''' <param name="sigdn">
        ''' A value from the <see cref="ShellItemGetDisplayName"/> enumeration that specifies the type of display name to retrieve.
        ''' </param>
        ''' 
        ''' <param name="refName">
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
        Friend Shared Function SHGetNameFromIDList(pidl As IntPtr,
                                                   sigdn As ShellItemGetDisplayName,
                                             <Out> ByRef refName As StringBuilder
        ) As HResult
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Creates and initializes a Shell item object from a parsing name.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/shobjidl_core/nf-shobjidl_core-shcreateitemfromparsingname"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="path">
        ''' A display name.
        ''' </param>
        '''
        ''' <param name="pbc">
        ''' Optional. A pointer to a bind context used to pass parameters as inputs and outputs to the parsing function. 
        ''' <para></para>
        ''' These passed parameters are often specific to the data source and are documented by the data source owners. 
        ''' <para></para>
        ''' For example, the file system data source accepts the name being parsed (as a <see cref="Win32FindDataW"/> structure), 
        ''' using the STR_FILE_SYS_BIND_DATA bind context parameter.
        ''' </param>
        ''' 
        ''' <param name="refIID">
        ''' A reference to the IID of the interface to retrieve through ppv, typically IID_IShellItem or IID_IShellItem2.
        ''' </param>
        ''' 
        ''' <param name="refShellItem">
        ''' When this method returns successfully, contains the interface pointer requested in <paramref name="refIID"/>. 
        ''' This is typically <see cref="IShellItem"/> or IShellItem2.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If this function succeeds, it returns <see cref="HResult.S_OK"/>. 
        ''' Otherwise, it returns an <see cref="HResult"/> error code.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("Shell32.dll", SetLastError:=True, CharSet:=CharSet.Auto, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function SHCreateItemFromParsingName(path As String,
                                                           pbc As IntPtr,
                                                           ByRef refIID As Guid,
                      <MarshalAs(UnmanagedType.Interface)> ByRef refShellItem As IShellItem
        ) As HResult
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Provides a default handler to extract an icon from a file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/bb762149%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="iconFile">
        ''' A pointer to a null-terminated buffer that contains the path and name of the file from which the icon is extracted.
        ''' </param>
        ''' 
        ''' <param name="iconIndex">
        ''' The location of the icon within the file named in pszIconFile.
        ''' <para></para>
        ''' If this is a positive number, it refers to the zero-based position of the icon in the file.
        ''' <para></para>
        ''' For instance, 0 refers to the 1st icon in the resource file and 2 refers to the 3rd.
        ''' If this is a negative number, it refers to the icon's resource ID.
        ''' </param>
        ''' 
        ''' <param name="flags">
        ''' A flag that controls the icon extraction.
        ''' </param>
        ''' 
        ''' <param name="refHiconLarge">
        ''' A pointer to an <c>HICON</c> that, when this function returns successfully, 
        ''' receives the handle of the large version of the icon specified in the LOWORD of nIconSize.
        ''' <para></para>
        ''' This value can be <see cref="IntPtr.Zero"/>.
        ''' </param>
        ''' 
        ''' <param name="refHiconSmall">
        ''' A pointer to an <c>HICON</c> that, when this function returns successfully, 
        ''' receives the handle of the small version of the icon specified in the <c>HIWORD</c> of <paramref name="iconSize"/> param.
        ''' <para></para>
        ''' This value can be <see cref="IntPtr.Zero"/>.
        ''' </param>
        ''' 
        ''' <param name="iconSize">
        ''' A value that contains the large icon size in its <c>LOWORD</c> and the small icon size in its <c>HIWORD</c>. 
        ''' <para></para>
        ''' Size is measured in pixels.
        ''' <para></para>
        ''' Pass <c>0</c> to specify default large and small sizes.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' This function can return one of these values:
        ''' <para></para>
        ''' <see cref="HResult.S_OK"/>, <see cref="HResult.S_FALSE"/> or <see cref="HResult.E_FAIL"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <SuppressUnmanagedCodeSecurity>
        <DllImport("Shell32.dll", SetLastError:=False, CharSet:=CharSet.Auto, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function SHDefExtractIcon(iconFile As String,
                                                iconIndex As Integer,
                                                flags As UInteger,
                                                ByRef refHiconLarge As IntPtr,
                                                ByRef refHiconSmall As IntPtr,
                                                iconSize As UInteger
        ) As HResult
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Displays a dialog box that allows the user to choose an icon from the 
        ''' selection available embedded in a resource such as an executable or DLL file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/shlobj_core/nf-shlobj_core-pickicondlg"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hWnd">
        ''' The handle of the parent window. This value can be <see cref="IntPtr.zero"/>.
        ''' </param>
        ''' 
        ''' <param name="iconPath">
        ''' A pointer to a string that contains the null-terminated, 
        ''' fully qualified path of the default resource that contains the icons. 
        ''' <para></para>
        ''' If the user chooses a different resource in the dialog, 
        ''' this buffer contains the path of that file when the function returns. 
        ''' <para></para>
        ''' This buffer should be at least MAX_PATH characters in length, or the returned path may be truncated. 
        ''' <para></para>
        ''' You should verify that the path is valid before using it.
        ''' </param>
        ''' 
        ''' <param name="iconPathSize">
        ''' The number of characters in <paramref name="iconPath"/>, including the terminating NULL character.
        ''' </param>
        ''' 
        ''' <param name="refIconIndex">
        ''' A pointer to an integer that on entry specifies the index of the initial selection and, 
        ''' when this function returns successfully, receives the index of the icon that was selected.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' Returns <see langword="true"/> if successful; otherwise, <see langword="false"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("Shell32.dll", SetLastError:=True, CharSet:=CharSet.Auto, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function PickIconDlg(hWnd As IntPtr,
                                           iconPath As String,
             <MarshalAs(UnmanagedType.U4)> iconPathSize As Integer,
                                           ByRef refIconIndex As Integer
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

#End Region

#Region " ShlwApi.dll "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Converts a numeric value into a string that represents the number expressed as a size value in bytes, 
        ''' kilobytes, megabytes, gigabytes, petabytes or exabytes, depending on the size.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/shlwapi/nf-shlwapi-strformatbytesize64a"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="number">
        ''' The numeric value to be converted.
        ''' </param>
        ''' 
        ''' <param name="buffer">
        ''' A pointer to a buffer that, when this function returns successfully, receives the converted number.
        ''' </param>
        ''' 
        ''' <param name="bufferSize">
        ''' The size of <paramref name="buffer"/>, in characters.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' Returns a pointer to the converted string, or NULL if the conversion fails.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("ShlwApi.dll", SetLastError:=True, ExactSpelling:=True, CharSet:=CharSet.Ansi)>
        Friend Shared Function StrFormatByteSize64A(number As ULong,
                                                    buffer As StringBuilder,
                                                    bufferSize As UInteger
        ) As IntPtr
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
        Friend Shared Function DestroyIcon(hIcon As IntPtr
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

#End Region

    End Class

End Namespace

#End Region
