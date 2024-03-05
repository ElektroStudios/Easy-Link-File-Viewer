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

#Region " msi.dll "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Examines an MSI shortcut file (*.lnk) and returns its product, feature name, and component if available.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://learn.microsoft.com/en-us/windows/win32/api/msi/nf-msi-msigetshortcuttargetw"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="shortcutPath">
        ''' A string specifying the full path to a shortcut file (*.lnk).
        ''' </param>
        ''' 
        ''' <param name="productCode">
        ''' A GUID for the product code of the shortcut. 
        ''' <para></para>
        ''' This string buffer must be 39 characters long. 
        ''' The first 38 characters are for the GUID, and the last character is for the terminating null character. 
        ''' <para></para>
        ''' This parameter can be null.
        ''' </param>
        ''' 
        ''' <param name="featureId">
        ''' The feature name of the shortcut.
        ''' <para></para>
        ''' The string buffer must be MAX_FEATURE_CHARS+1 characters long.
        ''' <para></para>
        ''' This parameter can be null.
        ''' </param>
        ''' 
        ''' <param name="componentCode">
        ''' A GUID of the component code. 
        ''' <para></para>
        ''' This string buffer must be 39 characters long. 
        ''' The first 38 characters are for the GUID, and the last character is for the terminating null character. 
        ''' <para></para>
        ''' This parameter can be null.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the function succeeds, the return value is 0.
        ''' <para></para>
        ''' If the function fails, the return value is non-zero.
        ''' <para></para>
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error()"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("msi.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Friend Shared Function MsiGetShortcutTargetW(shortcutPath As String,
                                                     productCode As StringBuilder,
                                                     featureId As StringBuilder,
                                                     componentCode As StringBuilder) As Integer
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Returns the full path to an installed component. 
        ''' <para></para>
        ''' If the key path for the component is a registry key then the registry key is returned.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://learn.microsoft.com/en-us/windows/win32/api/msi/nf-msi-msigetcomponentpathw"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="productCode">
        ''' A value that specifies an application's product code GUID. 
        ''' The function gets the path of installed components used by this application.
        ''' </param>
        ''' 
        ''' <param name="componentCode">
        ''' A value that specifies a component code GUID. 
        ''' The function gets the path of an installed component having this component code.
        ''' </param>
        ''' 
        ''' <param name="pathBuffer">
        ''' A buffer that eceives the path to the component. 
        ''' <para></para>
        ''' This parameter can be null. 
        ''' <para></para>
        ''' If the component is a registry key, the registry roots are represented numerically.
        ''' <para></para>
        ''' If this is a registry subkey path, there is a backslash at the end of the Key Path. 
        ''' <para></para>
        ''' If this is a registry value key path, there is no backslash at the end.
        ''' <para></para>
        ''' For example, a registry path on a 32-bit operating system of 'HKEY_CURRENT_USER\SOFTWARE\Microsoft' 
        ''' is returned as "01:\SOFTWARE\Microsoft".
        ''' </param>
        ''' 
        ''' <param name="refSizeBuffer">
        ''' specifies the size, in characters, of the buffer pointed to by the <paramref name="pathBuffer"/> parameter. 
        ''' <para></para>
        ''' On input, this is the full size of the buffer, including a space for a terminating null character. 
        ''' <para></para>
        ''' If the buffer passed in is too small, the count returned does not include the terminating null character.
        ''' <para></para>
        ''' If <paramref name="refSizeBuffer"/> is null, <paramref name="pathBuffer"/> can be null.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' Returns one of the following values:
        ''' <para></para>
        ''' INSTALLSTATE_NOTUSED (-7): The component being requested is disabled on the computer.
        ''' <para></para>
        ''' INSTALLSTATE_SOURCEABSENT (-4): The component source is inaccessible.
        ''' <para></para>
        ''' INSTALLSTATE_INVALIDARG (-2): One of the function parameters is invalid.
        ''' <para></para>
        ''' INSTALLSTATE_UNKNOWN (0): The product code or component ID is unknown. 
        ''' </returns>
        ''' INSTALLSTATE_ABSENT (2): The component is not installed.
        ''' <para></para>
        ''' INSTALLSTATE_LOCAL (3): The component is installed locally.
        ''' <para></para>
        ''' INSTALLSTATE_SOURCE (4): The component is installed to run from source.
        ''' <para></para>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("msi.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Friend Shared Function MsiGetComponentPathW(productCode As String,
                                                    componentCode As String,
                                              <Out> pathBuffer As StringBuilder,
                                  <[In], Out> ByRef refSizeBuffer As Integer) As Integer
        End Function

        ' Unused function. It may be helpful in the future.
#If DEBUG Then

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Returns the full path to an installed component. 
        ''' <para></para>
        ''' If the key path for the component is a registry key then the registry key is returned.
        ''' <para></para>
        ''' This function extends the existing <see cref="MsiGetComponentPathW"/> function 
        ''' to enable searches for components across user accounts and installation contexts.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://learn.microsoft.com/en-us/windows/win32/api/msi/nf-msi-msigetcomponentpathexw"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="productCode">
        ''' A value that specifies an application's product code GUID. 
        ''' The function gets the path of installed components used by this application.
        ''' </param>
        ''' 
        ''' <param name="componentCode">
        ''' A value that specifies a component code GUID. 
        ''' The function gets the path of an installed component having this component code.
        ''' </param>
        ''' 
        ''' <param name="userSid">
        ''' A string value that specifies the security identifier (SID) for a user in the system. 
        ''' <para></para>
        ''' The function gets the paths of installed components of applications installed under the user accounts 
        ''' identified by this SID. 
        ''' The special SID string s-1-1-0 (Everyone) specifies all users in the system. 
        ''' <para></para>
        ''' If this parameter is NULL, the function gets the path of an installed component for 
        ''' the currently logged-on user only.
        ''' </param>
        ''' 
        ''' <param name="context">
        ''' A flag that specifies the installation context. 
        ''' <para></para>
        ''' The function gets the paths of installed components of applications installed in 
        ''' the specified installation context. 
        ''' </param>
        ''' 
        ''' <param name="pathBuffer">
        ''' A buffer that eceives the path to the component. 
        ''' <para></para>
        ''' This parameter can be null. 
        ''' <para></para>
        ''' If the component is a registry key, the registry roots are represented numerically.
        ''' <para></para>
        ''' If this is a registry subkey path, there is a backslash at the end of the Key Path. 
        ''' <para></para>
        ''' If this is a registry value key path, there is no backslash at the end.
        ''' <para></para>
        ''' For example, a registry path on a 32-bit operating system of 'HKEY_CURRENT_USER\SOFTWARE\Microsoft' 
        ''' is returned as "01:\SOFTWARE\Microsoft".
        ''' </param>
        ''' 
        ''' <param name="refSizeBuffer">
        ''' specifies the size, in characters, of the buffer pointed to by the <paramref name="pathBuffer"/> parameter. 
        ''' <para></para>
        ''' On input, this is the full size of the buffer, including a space for a terminating null character. 
        ''' <para></para>
        ''' If the buffer passed in is too small, the count returned does not include the terminating null character.
        ''' <para></para>
        ''' If <paramref name="refSizeBuffer"/> is null, <paramref name="pathBuffer"/> can be null.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' Returns one of the following values:
        ''' <para></para>
        ''' INSTALLSTATE_NOTUSED (-7): The component being requested is disabled on the computer.
        ''' <para></para>
        ''' INSTALLSTATE_BADCONFIG (-6): Configuration data is corrupt.
        ''' <para></para>
        ''' INSTALLSTATE_SOURCEABSENT (-4): The component source is inaccessible.
        ''' <para></para>
        ''' INSTALLSTATE_INVALIDARG (-2): One of the function parameters is invalid.
        ''' <para></para>
        ''' INSTALLSTATE_UNKNOWN (-1): The product code or component ID is unknown.
        ''' <para></para>
        ''' INSTALLSTATE_BROKEN (0): The component is corrupt or partially missing in some way and requires repair. 
        ''' </returns>
        ''' INSTALLSTATE_ABSENT (2): The component is not installed.
        ''' <para></para>
        ''' INSTALLSTATE_LOCAL (3): The component is installed locally.
        ''' <para></para>
        ''' INSTALLSTATE_SOURCE (4): The component is installed to run from source.
        ''' <para></para>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("msi.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Friend Shared Function MsiGetComponentPathExW(productCode As String,
                                                      componentCode As String,
                                   <[In], [Optional]> userSid As String,
                                   <[In], [Optional]> context As MsiGetComponentPathContext,
                                    <Out, [Optional]> pathBuffer As StringBuilder,
                        <[In], Out, [Optional]> ByRef refSizeBuffer As Integer) As Integer
        End Function

#End If

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

        ' Unused functions. They may be helpful in the future.
#If DEBUG Then
        
        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Clones the first SHITEMID structure in an ITEMIDLIST structure.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/shlobj_core/nf-shlobj_core-ilclonefirst"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidl">
        ''' A pointer to the ITEMIDLIST structure that you want to clone.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A pointer to an ITEMIDLIST structure that contains the 
        ''' first SHITEMID structure from the ITEMIDLIST structure specified by <paramref name="pidl"/>. 
        ''' <para></para>
        ''' Returns NULL on failure.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("Shell32.dll", ExactSpelling:=True, SetLastError:=True)>
        Friend Shared Function ILCloneFirst(pidl As IntPtr
        ) As PIDL
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Clones an ITEMIDLIST structure.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/shlobj_core/nf-shlobj_core-ilclone"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidl">
        ''' A pointer to the ITEMIDLIST structure to be cloned.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' Returns a pointer to a copy of the ITEMIDLIST structure pointed to by <paramref name="pidl"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("Shell32.dll", EntryPoint:="ILClone", SetLastError:=True)>
        Friend Shared Function ILCloneIntPtr(pidl As IntPtr
        ) As IntPtr
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Combines two <c>ITEMIDLIST</c> structures.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/bb776437(v=vs.85).aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidlParent">
        ''' A pointer to the first <c>ITEMIDLIST</c> structure.
        ''' </param>
        ''' 
        ''' <param name="pidlChild">
        ''' A pointer to the second <c>ITEMIDLIST</c> structure. 
        ''' <para></para>
        ''' This structure is appended to the structure pointed to by <paramref name="pidlParent"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' Returns an <c>ITEMIDLIST</c> containing the combined structures. 
        ''' <para></para>
        ''' If you set either <paramref name="pidlParent"/> or <paramref name="pidlChild"/> to <see cref="IntPtr.Zero"/>, 
        ''' the returned <c>ITEMIDLIST</c> structure is a clone of the non-NULL parameter. 
        ''' <para></para>
        ''' Returns NULL if <paramref name="pidlParent"/> and <paramref name="pidlChild"/> are both set to <see cref="IntPtr.Zero"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <SuppressUnmanagedCodeSecurity>
        <DllImport("Shell32.dll", SetLastError:=True)>
        Friend Shared Function ILCombine(pidlParent As IntPtr,
                                         pidlChild As IntPtr
        ) As PIDL
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Combines two <c>ITEMIDLIST</c> structures.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/bb776437(v=vs.85).aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidlParent">
        ''' A pointer to the first <c>ITEMIDLIST</c> structure.
        ''' </param>
        ''' 
        ''' <param name="pidlChild">
        ''' A pointer to the second <c>ITEMIDLIST</c> structure. 
        ''' <para></para>
        ''' This structure is appended to the structure pointed to by <paramref name="pidlParent"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' Returns an <c>ITEMIDLIST</c> containing the combined structures. 
        ''' <para></para>
        ''' If you set either <paramref name="pidlParent"/> or <paramref name="pidlChild"/> to <see cref="IntPtr.Zero"/>, 
        ''' the returned <c>ITEMIDLIST</c> structure is a clone of the non-NULL parameter. 
        ''' <para></para>
        ''' Returns NULL if <paramref name="pidlParent"/> and <paramref name="pidlChild"/> are both set to <see cref="IntPtr.Zero"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("Shell32.dll", EntryPoint:="ILCombine", SetLastError:=True)>
        Friend Shared Function ILCombineIntPtr(pidlParent As IntPtr,
                                               pidlChild As IntPtr
        ) As IntPtr
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Returns the ITEMIDLIST structure associated with a specified file path.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/shlobj_core/nf-shlobj_core-ilcreatefrompath"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="path">
        ''' A pointer to a null-terminated Unicode string that contains the path. 
        ''' <para></para>
        ''' This string should be no more than MAX_PATH characters in length, including the terminating null character.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' Returns a pointer to an ITEMIDLIST structure that corresponds to the path.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("Shell32.dll", EntryPoint:="ILCreateFromPath", SetLastError:=True)>
        Friend Shared Function ILCreateFromPathIntPtr(<MarshalAs(UnmanagedType.LPWStr)> path As String
        ) As IntPtr
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Returns a pointer to the last SHITEMID structure in an ITEMIDLIST structure.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/shlobj_core/nf-shlobj_core-ilfindlastid"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidl">
        ''' A pointer to an ITEMIDLIST structure.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A pointer to the last SHITEMID structure in <paramref name="pidl"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("Shell32.dll", ExactSpelling:=True, SetLastError:=True)>
        Friend Shared Function ILFindLastID(pidl As IntPtr
        ) As IntPtr
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Frees an ITEMIDLIST structure allocated by the Shell.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/shlobj_core/nf-shlobj_core-ilfree"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidl">
        ''' A pointer to the ITEMIDLIST structure to be freed. 
        ''' <para></para>
        ''' This parameter can be NULL.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("Shell32.dll", ExactSpelling:=True, SetLastError:=True)>
        Friend Shared Sub ILFree(pidl As IntPtr)
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Returns the size, in bytes, of an ITEMIDLIST structure.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/shlobj_core/nf-shlobj_core-ilgetsize"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidl">
        ''' A pointer to an ITEMIDLIST structure.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The size of the ITEMIDLIST structure specified by <paramref name="pidl"/>, in bytes.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("Shell32.dll", ExactSpelling:=True, SetLastError:=True)>
        Friend Shared Function ILGetSize(pidl As IntPtr
        ) As UInteger
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Tests whether two ITEMIDLIST structures are equal in a binary comparison.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/shlobj_core/nf-shlobj_core-ilisequal"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidl1">
        ''' The first ITEMIDLIST structure.
        ''' </param>
        ''' 
        ''' <param name="pidl2">
        ''' The second ITEMIDLIST structure.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' Returns <see langword="True"/> if the two structures are equal, <see langword="False"/> otherwise.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("Shell32.dll", ExactSpelling:=True, SetLastError:=True)>
        Friend Shared Function ILIsEqual(pidl1 As IntPtr,
                                         pidl2 As IntPtr
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Tests whether an ITEMIDLIST structure is the parent of another ITEMIDLIST structure.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/shlobj_core/nf-shlobj_core-ilisparent"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidl1">
        ''' A pointer to an ITEMIDLIST (PIDL) structure that specifies the parent. This must be an absolute PIDL.
        ''' </param>
        ''' 
        ''' <param name="pidl2">
        ''' A pointer to an ITEMIDLIST (PIDL) structure that specifies the child. This must be an absolute PIDL.
        ''' </param>
        ''' 
        ''' <param name="immediate">
        ''' A Boolean value that is set to <see langword="True"/> to test for immediate parents of <paramref name="pidl2"/>, 
        ''' or <see langword="False"/> to test for any parents of <paramref name="pidl2"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' Returns <see langword="True"/> if <paramref name="pidl1"/> is a parent of <paramref name="pidl2"/>. 
        ''' <para></para>
        ''' If fImmediate is set to <see langword="True"/>, 
        ''' the function only returns <see langword="True"/> if <paramref name="pidl1"/> is the immediate parent of <paramref name="pidl2"/>. 
        ''' Otherwise, the function returns <see langword="False"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("Shell32.dll", ExactSpelling:=True, SetLastError:=True)>
        Friend Shared Function ILIsParent(pidl1 As IntPtr,
                                          pidl2 As IntPtr,
          <MarshalAs(UnmanagedType.Bool)> immediate As Boolean
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Removes the last SHITEMID structure from an ITEMIDLIST structure.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/shlobj_core/nf-shlobj_core-ilremovelastid"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidl">
        ''' A pointer to the ITEMIDLIST structure to be shortened. 
        ''' <para></para>
        ''' When the function returns, this variable points to the shortened structure.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' Returns <see langword="True"/> if successful, <see langword="False"/> otherwise.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("Shell32.dll", ExactSpelling:=True, SetLastError:=True)>
        Friend Shared Function ILRemoveLastID(pidl As IntPtr
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Clones an ITEMIDLIST structure.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/shlobj_core/nf-shlobj_core-ilclone"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidl">
        ''' A pointer to the ITEMIDLIST structure to be cloned.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' Returns a pointer to a copy of the ITEMIDLIST structure pointed to by <paramref name="pidl"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("Shell32.dll", ExactSpelling:=True, SetLastError:=True)>
        Friend Shared Function ILClone(pidl As IntPtr
        ) As PIDL
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Returns the ITEMIDLIST structure associated with a specified file path.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/shlobj_core/nf-shlobj_core-ilcreatefrompath"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="path">
        ''' A pointer to a null-terminated Unicode string that contains the path. 
        ''' <para></para>
        ''' This string should be no more than MAX_PATH characters in length, including the terminating null character.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' Returns a pointer to an ITEMIDLIST structure that corresponds to the path.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("Shell32.dll", CharSet:=CharSet.Auto, BestFitMapping:=False, ThrowOnUnmappableChar:=True, SetLastError:=True)>
        Friend Shared Function ILCreateFromPath(path As String
        ) As PIDL
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves the next SHITEMID structure in an ITEMIDLIST structure.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/shlobj_core/nf-shlobj_core-ilgetnext"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidl">
        ''' A pointer to a particular SHITEMID structure in a larger ITEMIDLIST structure.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' Returns a pointer to the SHITEMID structure that follows the one specified by <paramref name="pidl"/>. 
        ''' <para></para>
        ''' Returns <see cref="IntPtr.Zero"/> if <paramref name="pidl"/> points to the last SHITEMID structure.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("Shell32.dll", ExactSpelling:=True, SetLastError:=True)>
        Friend Shared Function ILGetNext(pidl As IntPtr
        ) As IntPtr
        End Function

#End If

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
