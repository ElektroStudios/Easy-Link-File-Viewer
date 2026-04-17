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
Imports DevCase.Win32.Common
Imports DevCase.Win32.Delegates.Delegates
Imports DevCase.Win32.Enums
Imports DevCase.Win32.Structures

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

#Region " dwmapi.dll "

        <DllImport("dwmapi.dll", PreserveSig:=True)>
        Friend Shared Function DwmSetWindowAttribute(hWnd As IntPtr,
                                                     attribute As Integer,
                                               ByRef refAttributeValue As Integer,
                                                     attributeSize As Integer
        ) As HResult
        End Function

#End Region

#Region " gdi32.dll "

        ''' <summary>
        ''' Computes the width and height of the specified string of text.
        ''' <para></para>
        ''' The <see cref="GetTextExtentPoint32"/> function uses the currently selected font to compute the dimensions of the string. 
        ''' The width and height, in logical units, are computed without considering any clipping.
        ''' <para></para>
        ''' Because some devices kern characters, the sum of the extents of the characters in 
        ''' a string may not be equal to the extent of the string.
        ''' </summary>
        ''' 
        ''' <remarks>
        ''' <see href="https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-gettextextentpoint32w"/>
        ''' </remarks>
        '''
        ''' <param name="hdc">
        ''' A handle to the device context.
        ''' </param>
        ''' 
        ''' <param name="text">
        ''' A buffer that specifies the text string. 
        ''' <para></para>
        ''' The string does not need to be null-terminated, because the <paramref name="textLength"/> parameter specifies the length of the string
        ''' </param>
        ''' 
        ''' <param name="textLength">
        ''' The length of the string pointed to by <paramref name="text"/> parameter.
        ''' </param>
        ''' 
        ''' <param name="refSize">
        ''' A <see cref="NativeSize"/> structure that receives the dimensions of the <paramref name="text"/> string, in logical units.
        ''' </param>
        ''' 
        ''' <returns>
        ''' If the function succeeds, the return value is <see langword="True"/>.
        ''' <para></para>
        ''' If the function fails, the return value is <see langword="False"/>. 
        ''' <para></para>
        ''' When this function returns the text extent, it assumes that the text is horizontal, that is, that the escapement is always 0. 
        ''' This is true for both the horizontal and vertical measurements of the text. 
        ''' </returns>
        <DllImport("gdi32.dll", CharSet:=CharSet.Auto, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function GetTextExtentPoint32(hdc As IntPtr,
                                                     text As String,
                                                     textLength As Integer,
                                               ByRef refSize As NativeSize
        ) As Boolean
        End Function

        ''' <summary>
        ''' Creates a logical brush that has the specified solid color.
        ''' <para></para>
        ''' A solid brush is a bitmap that the system uses to paint the interiors of filled shapes.
        ''' <para></para>
        ''' After an application creates a brush by calling <see cref="CreateSolidBrush"/>,
        ''' it can select that brush into any device context by calling the <see cref="NativeMethods.SelectObject"/> function.
        ''' <para></para>
        ''' When you no longer need the brush, call the <see cref="NativeMethods.DeleteObject"/> function to delete it.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/dd183518%28v=vs.85%29.aspx"/>
        ''' </remarks>
        '''
        ''' <param name="color">
        ''' The color of the brush.
        ''' <para></para>
        ''' To create a <c>COLORREF</c> color value, use the <c>RGB</c> macro.
        ''' </param>
        '''
        ''' <returns>
        ''' If the function succeeds, the return value identifies a logical brush.
        ''' <para></para>
        ''' If the function fails, the return value is <see cref="IntPtr.Zero"/>.
        ''' </returns>
        <SuppressUnmanagedCodeSecurity>
        <DllImport("gdi32.dll", SetLastError:=True)>
        Friend Shared Function CreateSolidBrush(color As NativeColor
        ) As IntPtr
        End Function

        ''' <summary>
        ''' Sets the current background color to the specified color value, 
        ''' or to the nearest physical color if the device cannot represent the specified color value.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-setbkcolor"/>
        ''' </remarks>
        '''
        ''' <param name="hdc">
        ''' A handle to the device context.
        ''' </param>
        ''' 
        ''' <param name="color">
        ''' The new background color.
        ''' </param>
        '''
        ''' <returns>
        ''' If the function succeeds, the return value specifies the previous background color as a COLORREF value.
        ''' <para></para>
        ''' If the function fails, the return value is CLR_INVALID (-1 or 0xFFFFFFFF).
        ''' </returns>
        <DllImport("gdi32.dll")>
        Friend Shared Function SetBkColor(hdc As IntPtr,
                                   color As NativeColor
        ) As NativeColor
        End Function

        ''' <summary>
        ''' Sets the text color for the specified device context to the specified color.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-settextcolor"/>
        ''' </remarks>
        '''
        ''' <param name="hdc">
        ''' A handle to the device context.
        ''' </param>
        ''' 
        ''' <param name="color">
        ''' The new color for the text.
        ''' </param>
        '''
        ''' <returns>
        ''' If the function succeeds, the return value is a color reference for the previous text color as a COLORREF value.
        ''' <para></para>
        ''' If the function fails, the return value is CLR_INVALID (-1 or 0xFFFFFFFF).
        ''' </returns>
        <DllImport("gdi32.dll")>
        Friend Shared Function SetTextColor(hdc As IntPtr,
                                     color As Integer
        ) As NativeColor
        End Function

        ''' <summary>
        ''' Sts the background mix mode of the specified device context (DC). 
        ''' <para></para>
        ''' The background mix mode is used with text, hatched brushes, and pen styles that are not solid lines.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/wingdi/nf-wingdi-setbkmode"/>
        ''' </remarks>
        '''
        ''' <param name="hdc">
        ''' A handle to the device context.
        ''' </param>
        ''' 
        ''' <param name="mode">
        ''' The background mode.
        ''' </param>
        '''
        ''' <returns>
        ''' If the function succeeds, the return value specifies the previous background mode. 
        ''' <para></para>
        ''' If the function fails, the return value is <see cref="BackgroundMode.Error"/>.
        ''' </returns>
        <DllImport("gdi32.dll", ExactSpelling:=True, SetLastError:=True)>
        Friend Shared Function SetBkMode(hdc As IntPtr,
                                  mode As BackgroundMode
        ) As BackgroundMode
        End Function

        ''' <summary>
        ''' Deletes a logical pen, brush, font, bitmap, region, or palette,
        ''' freeing all system resources associated with the object.
        ''' <para></para>
        ''' After the object is deleted, the specified handle is no longer valid.
        ''' <para></para>
        ''' Do not delete a drawing object (pen or brush) while it is still selected into a DC.
        ''' <para></para>
        ''' When a pattern brush is deleted, the bitmap associated with the brush is not deleted. 
        ''' The bitmap must be deleted independently.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms633540%28v=vs.85%29.aspx"/>
        ''' </remarks>
        '''
        ''' <param name="hObject">
        ''' A handle to a logical pen, brush, font, bitmap, region, or palette.
        ''' </param>
        '''
        ''' <returns>
        ''' If the function succeeds, the return value is <see langword="True"/>.
        ''' <para></para>
        ''' If the specified handle is not valid or is currently selected into a DC, the return value is <see langword="False"/>.
        ''' </returns>
        <SuppressUnmanagedCodeSecurity>
        <DllImport("gdi32.dll", CharSet:=CharSet.Auto, ExactSpelling:=False, SetLastError:=True)>
        Friend Shared Function DeleteObject(hObject As IntPtr
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

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

        ''' <summary>
        ''' Gets the thread identifier of the calling thread.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms683183%28v=vs.85%29.aspx"/>
        ''' </remarks>
        '''
        ''' <returns>
        ''' The thread identifier of the calling thread.
        ''' </returns>
        <SuppressUnmanagedCodeSecurity>
        <DllImport("Kernel32.dll", SetLastError:=False)>
        Friend Shared Function GetCurrentThreadId(
        ) As UInteger
        End Function

        ''' <summary>
        ''' Retrieves the window handle used by the console associated with the calling process.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms683175%28v=vs.85%29.aspx"/>
        ''' </remarks>
        '''
        ''' <returns>
        ''' The return value is a handle to the window used by the console associated with the calling process
        ''' or <see cref="IntPtr.Zero"/> if there is no such associated console.
        ''' </returns>
        <SuppressUnmanagedCodeSecurity>
        <DllImport("Kernel32.dll", SetLastError:=False)>
        Friend Shared Function GetConsoleWindow() As IntPtr
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

        ''' <summary>
        ''' Changes the size, position, and Z order of a child, pop-up, or top-level window.
        ''' <para></para>
        ''' These windows are ordered according to their appearance on the screen.
        ''' <para></para>
        ''' The topmost window receives the highest rank and is the first window in the Z order.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms633545(v=vs.85).aspx"/>
        ''' </remarks>
        '''
        ''' <param name="hWnd">
        ''' A handle to the window.
        ''' </param>
        ''' 
        ''' <param name="hWndInsertAfter">
        ''' A handle to the window to precede the positioned window in the Z order.
        ''' </param>
        ''' 
        ''' <param name="x">
        ''' The new position of the left side of the window, in client coordinates.
        ''' </param>
        ''' 
        ''' <param name="y">
        ''' The new position of the top of the window, in client coordinates.
        ''' </param>
        ''' 
        ''' <param name="cx">
        ''' The new width of the window, in pixels.
        ''' </param>
        ''' 
        ''' <param name="cy">
        ''' The new height of the window, in pixels.
        ''' </param>
        ''' 
        ''' <param name="uFlags">
        ''' The window sizing and positioning flags.
        ''' </param>
        '''
        ''' <returns>
        ''' If the function succeeds, the return value is <see langword="True"/>.
        ''' <para></para>
        ''' If the function fails, the return value is <see langword="False"/>.
        ''' <para></para>
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>. 
        ''' </returns>
        <SuppressUnmanagedCodeSecurity>
        <DllImport("User32.dll", SetLastError:=True)>
        Friend Shared Function SetWindowPos(hWnd As IntPtr,
                                            hWndInsertAfter As IntPtr,
                                            x As Integer,
                                            y As Integer,
                                            cx As Integer,
                                            cy As Integer,
                                            uFlags As SetWindowPosFlags
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Releases a device context (DC), freeing it for use by other applications.
        ''' <para></para>
        ''' The effect of the <see cref="ReleaseDC"/> function depends on the type of DC. 
        ''' It frees only common and window DCs. 
        ''' It has no effect on class or private DCs.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/dd162920%28v=vs.85%29.aspx"/>
        ''' </remarks>
        '''
        ''' <param name="hWnd">
        ''' A <see cref="IntPtr"/> handle to the window whose DC is to be released.
        ''' </param>
        ''' 
        ''' <param name="hdc">
        ''' A <see cref="IntPtr"/> handle to the DC to be released.
        ''' </param>
        '''
        ''' <returns>
        ''' <see langword="True"/> if the DC was released, <see langword="False"/> if the DC was not released.
        ''' </returns>
        <SuppressUnmanagedCodeSecurity>
        <DllImport("User32.dll", SetLastError:=False)>
        Friend Shared Function ReleaseDC(hWnd As IntPtr,
                                         hdc As IntPtr
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function


        ''' <summary>
        ''' Retrieves the coordinates of a window's client area.
        ''' <para></para>
        ''' The client coordinates specify the upper-left and lower-right corners of the client area.
        ''' <para></para>
        ''' Because client coordinates are relative to the upper-left corner of a window's client area, 
        ''' the coordinates of the upper-left corner are (0,0). 
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms633503%28v=vs.85%29.aspx"/>
        ''' </remarks>
        '''
        ''' <param name="hWnd">
        ''' A handle to the window whose client coordinates are to be retrieved. 
        ''' </param>
        ''' 
        ''' <param name="refRect">
        ''' A <see cref="NativeRectangle"/> structure that receives the client coordinates.
        ''' <para></para>
        ''' The left and top members are zero. The right and bottom members contain the width and height of the window.
        ''' </param>
        '''
        ''' <returns>
        ''' <see langword="True"/> if the function succeeds, <see langword="False"/> otherwise.
        ''' </returns>
        <SuppressUnmanagedCodeSecurity>
        <DllImport("User32.dll", SetLastError:=True)>
        Friend Shared Function GetClientRect(hWnd As IntPtr,
                                 <[In], Out> ByRef refRect As NativeRectangle
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Creates an overlapped, pop-up, or child window with an extended window style. 
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/winuser/nf-winuser-createwindowexa"/>
        ''' </remarks>
        '''
        ''' <param name="exStyle">
        ''' The extended window style of the window being created.
        ''' </param>
        ''' 
        ''' <param name="className">
        ''' A null-terminated string that specifies the window class name. 
        ''' <para></para>
        ''' The class name can be any name registered with <see cref="RegisterClass"/> or <see cref="RegisterClassEx"/>, 
        ''' provided that the module that registers the class is also the module that creates the window. 
        ''' <para></para>
        ''' The class name can also be any of the predefined system class names.
        ''' </param>
        ''' 
        ''' <param name="windowName">
        ''' The window name. 
        ''' <para></para>
        ''' If the window style specifies a title bar, the window title pointed to by lpWindowName is displayed in the title bar. 
        ''' <para></para>
        ''' When using <see cref="NativeMethods.CreateWindowEx"/> to create controls, such as buttons, check boxes, and static controls, 
        ''' use <paramref name="windowName"/> to specify the text of the control. 
        ''' <para></para>
        ''' When creating a static control with the SS_ICON style, use <paramref name="windowName"/> to specify the icon name or identifier. 
        ''' <para></para>
        ''' To specify an identifier, use the syntax "#num".
        ''' </param>
        ''' 
        ''' <param name="style">
        ''' The style of the window being created. 
        ''' </param>
        ''' 
        ''' <param name="x">
        ''' The initial horizontal position of the window. 
        ''' <para></para>
        ''' For an overlapped or pop-up window, the x parameter is the initial x-coordinate of the window's upper-left corner, in screen coordinates. 
        ''' <para></para>
        ''' For a child window, x is the x-coordinate of the upper-left corner of the window relative to the upper-left corner of the 
        ''' parent window's client area. 
        ''' <para></para>
        ''' If x is set to CW_USEDEFAULT, the system selects the default position for the window's upper-left corner and ignores the y parameter. 
        ''' <para></para>
        ''' CW_USEDEFAULT is valid only for overlapped windows; 
        ''' if it is specified for a pop-up or child window, the x and y parameters are set to zero.
        ''' </param>
        ''' 
        ''' <param name="y">
        ''' The initial vertical position of the window. 
        ''' <para></para>
        ''' For an overlapped or pop-up window, the y parameter is the initial y-coordinate of the window's upper-left corner, in screen coordinates. 
        ''' <para></para>
        ''' For a child window, y is the initial y-coordinate of the upper-left corner of the child window relative to the upper-left corner of the 
        ''' parent window's client area. 
        ''' <para></para>
        ''' For a list box y is the initial y-coordinate of the upper-left corner of the list box's client area relative to the 
        ''' upper-left corner of the parent window's client area.
        ''' <para></para>
        ''' If an overlapped window is created with the WS_VISIBLE style bit set and the x parameter is set to CW_USEDEFAULT, 
        ''' then the y parameter determines how the window is shown. 
        ''' <para></para>
        ''' If the y parameter is CW_USEDEFAULT, then the window manager calls ShowWindow with the SW_SHOW flag after the window has been created. 
        ''' <para></para>
        ''' If the y parameter is some other value, then the window manager calls 
        ''' <see cref="NativeMethods.ShowWindow"/> with that value as the <c>windowState</c> (<c>nCmdShow</c>) parameter.
        ''' </param>
        ''' 
        ''' <param name="width">
        ''' The width, in device units, of the window. 
        ''' <para></para>
        ''' For overlapped windows, nWidth is the window's width, in screen coordinates, or CW_USEDEFAULT. 
        ''' <para></para>
        ''' If <paramref name="width"/> is CW_USEDEFAULT, the system selects a default width and height for the window; 
        ''' the default width extends from the initial x-coordinates to the right edge of the screen; 
        ''' the default height extends from the initial y-coordinate to the top of the icon area. 
        ''' <para></para>
        ''' CW_USEDEFAULT is valid only for overlapped windows; 
        ''' if CW_USEDEFAULT is specified for a pop-up or child window, the nWidth and nHeight parameter are set to zero.
        ''' </param>
        ''' 
        ''' <param name="height">
        ''' The height, in device units, of the window. 
        ''' <para></para>
        ''' For overlapped windows, nHeight is the window's height, in screen coordinates. 
        ''' <para></para>
        ''' If the <paramref name="width"/> parameter is set to CW_USEDEFAULT, the system ignores nHeight.
        ''' </param>
        ''' 
        ''' <param name="hWndParent">
        ''' A handle to the parent or owner window of the window being created. 
        ''' <para></para>
        ''' To create a child window or an owned window, supply a valid window handle. 
        ''' <para></para>
        ''' This parameter is optional for pop-up windows.
        ''' <para></para>
        ''' To create a message-only window, supply HWND_MESSAGE or a handle to an existing message-only window.
        ''' </param>
        ''' 
        ''' <param name="hMenu">
        ''' A handle to a menu, or specifies a child-window identifier, depending on the window style. 
        ''' <para></para>
        ''' For an overlapped or pop-up window, <paramref name="hMenu"/> identifies the menu to be used with the window; 
        ''' it can be NULL if the class menu is to be used. 
        ''' <para></para>
        ''' For a child window, <paramref name="hMenu"/> specifies the child-window identifier, 
        ''' an integer value used by a dialog box control to notify its parent about events. 
        ''' <para></para>
        ''' The application determines the child-window identifier; it must be unique for all child windows with the same parent window.
        ''' </param>
        ''' 
        ''' <param name="hInstance">
        ''' A handle to the instance of the module to be associated with the window.
        ''' </param>
        ''' 
        ''' <param name="param">
        ''' Pointer to a value to be passed to the window through the CREATESTRUCT structure (lpCreateParams member) 
        ''' pointed to by the lParam param of the <see cref="WindowMessages.WM_CREATE"/> message. 
        ''' This message is sent to the created window by this function before it returns.
        ''' <para></para>
        ''' If an application calls <see cref="NativeMethods.CreateWindowEx"/> to create a MDI client window, 
        ''' <paramref name="param"/> should point to a CLIENTCREATESTRUCT structure. 
        ''' <para></para>
        ''' If an MDI client window calls <see cref="NativeMethods.CreateWindowEx"/> to create an MDI child window,
        ''' <paramref name="param"/> should point to a MDICREATESTRUCT structure. 
        ''' <para></para>
        ''' <paramref name="param"/> may be <see cref="IntPtr.Zero"/> if no additional data is needed.
        ''' </param>
        '''
        ''' <returns>
        ''' If the function succeeds, the return value is a handle to the new window.
        ''' <para></para>
        ''' If the function fails, the return value is <see cref="IntPtr.Zero"/>.
        ''' </returns>
        <DllImport("User32.dll", SetLastError:=True, CharSet:=CharSet.Auto, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function CreateWindowEx(exStyle As WindowStylesEx,
                                               className As String,
                                               windowName As String,
                                               style As WindowStyles,
                                               x As Integer,
                                               y As Integer,
                                               width As Integer,
                                               height As Integer,
                                  <[Optional]> hWndParent As IntPtr,
                                  <[Optional]> hMenu As IntPtr,
                                               hInstance As IntPtr,
                                  <[Optional]> param As IntPtr
        ) As IntPtr
        End Function

        ''' <summary>
        ''' Passes message information to the specified window procedure.
        ''' </summary>
        ''' 
        ''' <remarks>
        ''' <see href="https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-callwindowprocw"/>
        ''' </remarks>
        '''
        ''' <param name="prevWndFunc">
        ''' The previous window procedure. 
        ''' <para></para>
        ''' If this value is obtained by calling the <see cref="NativeMethods.GetWindowLong"/> function with the 
        ''' flags parameter set to <see cref="WindowLongValues.WndProc"/> or <see cref="WindowLongValues.DlgProc"/>, 
        ''' it is actually either the address of a window or dialog box procedure, 
        ''' or a special internal value meaningful only to <see cref="NativeMethods.CallWindowProc"/>.
        ''' </param>
        ''' 
        ''' <param name="hWnd">
        ''' A handle to the window procedure to receive the message.
        ''' </param>
        ''' 
        ''' <param name="msg">
        ''' The message.
        ''' </param>
        ''' 
        ''' <param name="wParam">
        ''' Additional message-specific information. 
        ''' <para></para>
        ''' The contents of this parameter depend on the value of the <paramref name="msg"/> parameter.
        ''' </param>
        ''' 
        ''' <param name="lParam">
        ''' Additional message-specific information. 
        ''' <para></para>
        ''' The contents of this parameter depend on the value of the <paramref name="msg"/> parameter.
        ''' </param>
        ''' 
        ''' <returns>
        ''' The return value specifies the result of the message processing and depends on the message sent.
        ''' </returns>
        <DllImport("User32.dll", SetLastError:=False, CharSet:=CharSet.Auto)>
        Friend Shared Function CallWindowProc(prevWndFunc As WndProc,
                                       hWnd As IntPtr,
                                       msg As UInteger,
                                       wParam As IntPtr,
                                       lParam As IntPtr
        ) As IntPtr
        End Function

        ''' <summary>
        ''' Passes message information to the specified window procedure.
        ''' </summary>
        ''' 
        ''' <remarks>
        ''' <see href="https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-callwindowprocw"/>
        ''' </remarks>
        '''
        ''' <param name="prevWndFunc">
        ''' The previous window procedure. 
        ''' <para></para>
        ''' If this value is obtained by calling the <see cref="NativeMethods.GetWindowLong"/> function with the 
        ''' flags parameter set to <see cref="WindowLongValues.WndProc"/> or <see cref="WindowLongValues.DlgProc"/>, 
        ''' it is actually either the address of a window or dialog box procedure, 
        ''' or a special internal value meaningful only to <see cref="NativeMethods.CallWindowProc"/>.
        ''' </param>
        ''' 
        ''' <param name="hWnd">
        ''' A handle to the window procedure to receive the message.
        ''' </param>
        ''' 
        ''' <param name="msg">
        ''' The message.
        ''' </param>
        ''' 
        ''' <param name="wParam">
        ''' Additional message-specific information. 
        ''' <para></para>
        ''' The contents of this parameter depend on the value of the <paramref name="msg"/> parameter.
        ''' </param>
        ''' 
        ''' <param name="lParam">
        ''' Additional message-specific information. 
        ''' <para></para>
        ''' The contents of this parameter depend on the value of the <paramref name="msg"/> parameter.
        ''' </param>
        ''' 
        ''' <returns>
        ''' The return value specifies the result of the message processing and depends on the message sent.
        ''' </returns>
        <DllImport("User32.dll", SetLastError:=False, CharSet:=CharSet.Auto)>
        Friend Shared Function CallWindowProc(<[In]> prevWndFunc As IntPtr,
                                              hWnd As IntPtr,
                                              msg As UInteger,
                                              wParam As IntPtr,
                                              lParam As IntPtr
        ) As IntPtr
        End Function

        ''' <summary>
        ''' Destroys the specified window.
        ''' The function sends <see cref="WindowMessages.WM_Destroy"/> and <see cref="WindowMessages.WM_NcDestroy"/> messages to the window 
        ''' to deactivate it and remove the keyboard focus from it.
        ''' <para></para>
        ''' The function also destroys the window's menu, flushes the thread message queue, destroys timers, 
        ''' removes clipboard ownership, and breaks the clipboard viewer chain (if the window is at the top of the viewer chain).
        ''' <para></para>
        ''' If the specified window is a parent or owner window, 
        ''' <see cref="NativeMethods.DestroyWindow"/> automatically destroys the associated child or owned windows when 
        ''' it destroys the parent or owner window.
        ''' <para></para>
        ''' The function first destroys child or owned windows, and then it destroys the parent or owner window.
        ''' <para></para>
        ''' <see cref="NativeMethods.DestroyWindow"/> also destroys modeless dialog boxes created by the <c>CreateDialog</c> function.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms632682%28v=vs.85%29.aspx"/>
        ''' </remarks>
        '''
        ''' <param name="hWnd">
        ''' A handle to the window to be destroyed. 
        ''' </param>
        '''
        ''' <returns>
        ''' If the function succeeds, the return value is <see cref="IntPtr.Zero"/>.
        ''' <para></para>
        ''' If the function fails, the return value is equal to a handle to the local memory object.
        ''' <para></para>
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>. 
        ''' </returns>
        <SuppressUnmanagedCodeSecurity>
        <DllImport("User32.dll", EntryPoint:="DestroyWindow", SetLastError:=True,
                   CallingConvention:=CallingConvention.StdCall)>
        Friend Shared Function DestroyWindow(hWnd As IntPtr
        ) As IntPtr
        End Function

        ''' <summary>
        ''' Destroys a modal dialog box, causing the system to end any processing for the dialog box.
        ''' </summary>
        ''' 
        ''' <remarks>
        ''' <see href="https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-enddialog"/>
        ''' </remarks>
        '''
        ''' <param name="hDlg">
        ''' A handle to the dialog box to be destroyed.
        ''' </param>
        ''' 
        ''' <param name="result">
        ''' The value to be returned to the application from the function that created the dialog box.
        ''' </param>
        ''' 
        ''' <returns>
        ''' If the function succeeds, the return value is <see langword="True"/>.
        ''' <para></para>
        ''' If the function fails, the return value is <see langword="False"/>.
        ''' <para></para>
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>.
        ''' </returns>
        <DllImport("User32.dll", SetLastError:=True, ExactSpelling:=True)>
        Friend Shared Function EndDialog(hDlg As IntPtr,
                                  result As IntPtr
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Changes the text of the specified window's title bar (if it has one). 
        ''' <para></para>
        ''' If the specified window is a control, the text of the control is changed. 
        ''' <para></para>
        ''' However, <see cref="NativeMethods.SetWindowText"/> cannot change the text of a control in another application.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms633546(v=vs.85).aspx"/>
        ''' </remarks>
        '''
        ''' <param name="hWnd">
        ''' A handle to the window or control whose text is to be changed. 
        ''' </param>
        ''' 
        ''' <param name="text">
        ''' The new title or control text. 
        ''' </param>
        '''
        ''' <returns>
        ''' If the function succeeds, the return value is <see langword="True"/>
        ''' <para></para>
        ''' If the function fails, the return value is <see langword="False"/>. 
        ''' <para></para>
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>.
        ''' </returns>
        <SuppressUnmanagedCodeSecurity>
        <DllImport("User32.dll", CharSet:=CharSet.Auto, SetLastError:=True, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function SetWindowText(hWnd As IntPtr,
                                             text As String
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Retrieves a <see cref="IntPtr"/> handle to a device context (DC) for the client area of a specified window or for the entire screen.
        ''' <para></para>
        ''' You can use the returned handle in subsequent GDI functions to draw in the DC.
        ''' <para></para>
        ''' The device context is an opaque data structure, whose values are used internally by GDI.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/dd144871%28v=vs.85%29.aspx"/>
        ''' </remarks>
        '''
        ''' <param name="hWnd">
        ''' A <see cref="IntPtr"/> handle to the window whose DC is to be retrieved. 
        ''' <para></para>
        ''' If this value is <see cref="IntPtr.Zero"/>, <see cref="GetDC"/> retrieves the DC for the entire screen.
        ''' </param>
        '''
        ''' <returns>
        ''' If the function succeeds, the return value is a <see cref="IntPtr"/> handle to the DC for the specified window's client area.
        ''' <para></para>
        ''' If the function fails, the return value is <see cref="IntPtr.Zero"/>.
        ''' </returns>
        <SuppressUnmanagedCodeSecurity>
        <DllImport("User32.dll", SetLastError:=False)>
        Friend Shared Function GetDC(hWnd As IntPtr
        ) As IntPtr
        End Function

        ''' <summary>
        ''' Activates a window. The window must be attached to the calling thread's message queue.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms646311%28v=vs.85%29.aspx"/>
        ''' </remarks>
        '''
        ''' <param name="hWnd">
        ''' A <see cref="IntPtr"/> handle to the top-level window to be activated.
        ''' </param>
        '''
        ''' <returns>
        ''' If the function succeeds, the return value is the <see cref="IntPtr"/> handle to the window that was previously active.
        ''' <para></para>
        ''' If the function fails, the return value is <see cref="IntPtr.Zero"/>.
        ''' </returns>
        <SuppressUnmanagedCodeSecurity>
        <DllImport("User32.dll", SetLastError:=True)>
        Friend Shared Function SetActiveWindow(hWnd As IntPtr
        ) As IntPtr
        End Function

        ''' <summary>
        ''' Retrieves the dimensions of the bounding rectangle of the specified window. 
        ''' <para></para>
        ''' The dimensions are given in screen coordinates that are relative to the upper-left corner of the screen.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms633519%28v=vs.85%29.aspx"/>
        ''' </remarks>
        '''
        ''' <param name="hWnd">
        ''' A <see cref="IntPtr"/> handle to the window.
        ''' </param>
        ''' 
        ''' <param name="refRect">
        ''' A <see cref="NativeRectangle"/> structure that receives the screen coordinates of the 
        ''' upper-left and lower-right corners of the window.
        ''' </param>
        '''
        ''' <returns>
        ''' <see langword="True"/> if the function succeeds, <see langword="False"/> otherwise.
        ''' </returns>
        <SuppressUnmanagedCodeSecurity>
        <DllImport("User32.dll", SetLastError:=True)>
        Friend Shared Function GetWindowRect(hWnd As IntPtr,
                                       <Out> ByRef refRect As NativeRectangle
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Retrieves the name of the class to which the specified window belongs.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms633582(v=vs.85).aspx"/>
        ''' </remarks>
        '''
        ''' <param name="hWnd">
        ''' A handle to the window and, indirectly, the class to which the window belongs.
        ''' </param>
        ''' 
        ''' <param name="className">
        ''' The class name string. 
        ''' </param>
        ''' 
        ''' <param name="maxCount">
        ''' The length of the <paramref name="className"/> buffer, in characters. 
        ''' <para></para>
        ''' The buffer must be large enough to include the terminating null character; 
        ''' otherwise, the class name string is truncated to <paramref name="maxCount"/>-1 characters. 
        ''' </param>
        '''
        ''' <returns>
        ''' If the function succeeds, the return value is the number of characters copied to the buffer, 
        ''' not including the terminating null character.
        ''' <para></para>
        ''' If the function fails, the return value is <c>0</c>. 
        ''' <para></para>
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>.
        ''' </returns>
        <SuppressUnmanagedCodeSecurity>
        <DllImport("User32.dll", CharSet:=CharSet.Auto, SetLastError:=True, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function GetClassName(hWnd As IntPtr,
                                            className As StringBuilder,
                                            maxCount As Integer
        ) As Integer
        End Function

        ''' <summary>
        ''' Retrieves a handle to a control in the specified dialog box.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms645481%28v=vs.85%29.aspx"/>
        ''' </remarks>
        '''
        ''' <param name="hWnd">
        ''' A handle to the dialog box that contains the control
        ''' .</param>
        ''' 
        ''' <param name="index">
        ''' The index of the item to be retrieved.
        ''' </param>
        '''
        ''' <returns>
        ''' If the function succeeds, the return value is the window handle of the specified control.
        ''' <para></para>
        ''' If the function fails, the return value is <see cref="IntPtr.Zero"/>, 
        ''' indicating an invalid dialog box handle or a nonexistent control.
        ''' </returns>
        <SuppressUnmanagedCodeSecurity>
        <DllImport("User32.dll", SetLastError:=False)>
        Friend Shared Function GetDlgItem(hWnd As IntPtr,
                                          index As Integer
        ) As IntPtr
        End Function

        ''' <summary>
        ''' Determines the visibility state of the specified window.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms633530%28v=vs.85%29.aspx"/>
        ''' </remarks>
        '''
        ''' <param name="hWnd">
        ''' A handle to the window to be tested. 
        ''' </param>
        '''
        ''' <returns>
        ''' If the specified window, its parent window, its parent's parent window, and so forth, 
        ''' have the <see cref="WindowStyles.Visible"/> style, the return value is <see langword="True"/>.
        ''' <para></para>
        ''' Otherwise, the return value is <see langword="False"/>.
        ''' <para></para>
        ''' Because the return value specifies whether the window has the WS_VISIBLE style, 
        ''' it may be nonzero even if the window is totally obscured by other windows.
        ''' </returns>
        <SuppressUnmanagedCodeSecurity>
        <DllImport("User32.dll", SetLastError:=True)>
        Friend Shared Function IsWindowVisible(hWnd As IntPtr
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Enumerates all nonchild windows associated with a thread by passing the handle to each window, 
        ''' in turn, to an application-defined callback function.
        ''' <para></para>
        ''' <see cref="EnumThreadWindows"/> continues until the last window is enumerated 
        ''' or the callback function returns <see langword="False"/>.
        ''' <para></para>
        ''' To enumerate child windows of a particular window, use the <see cref="EnumChildWindows"/> function.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms633495%28v=vs.85%29.aspx"/>
        ''' </remarks>
        '''
        ''' <param name="threadId">
        ''' The identifier of the thread whose windows are to be enumerated.
        ''' </param>
        ''' 
        ''' <param name="lpEnumFunc">
        ''' An application-defined callback function.
        ''' </param>
        ''' 
        ''' <param name="lParam">
        ''' An application-defined value to be passed to the callback function. 
        ''' </param>
        '''
        ''' <returns>
        ''' If the callback function returns <see langword="True"/> for all windows in the thread specified by threadId, 
        ''' the return value is <see langword="True"/>.
        ''' <para></para>
        ''' If the callback function returns <see langword="False"/> on any enumerated window, 
        ''' or if there are no windows found in the thread specified by threadId, 
        ''' the return value is <see langword="False"/>.
        ''' </returns>
        <SuppressUnmanagedCodeSecurity>
        <DllImport("User32.dll", SetLastError:=True)>
        Friend Shared Function EnumThreadWindows(threadId As UInteger,
                                                 lpEnumFunc As EnumThreadWindowsProc,
                                                 lParam As IntPtr
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Changes an attribute of the specified window.
        ''' <para></para>
        ''' The function also sets the 32-bit (<c>LONG</c>) value at the specified offset into the extra window memory.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms633591%28v=vs.85%29.aspx"/>
        ''' </remarks>
        '''
        ''' <param name="hWnd">
        ''' A handle to the window and, indirectly, the class to which the window belongs.
        ''' </param>
        ''' 
        ''' <param name="flags">
        ''' The zero-based offset to the value to be set. 
        ''' <para></para>
        ''' Valid values are in the range zero through the number of bytes of extra window memory, 
        ''' minus the size of an integer.
        ''' </param>
        ''' 
        ''' <param name="newLong">
        ''' The replacement value.
        ''' </param>
        '''
        ''' <returns>
        ''' If the function succeeds, the return value is the previous value of the specified 32-bit integer.
        ''' <para></para>
        ''' If the function fails, the return value is zero.
        ''' <para></para>
        ''' If the previous value of the specified 32-bit integer is zero, 
        ''' and the function succeeds, the return value is zero.
        ''' <para></para>
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>. 
        ''' </returns>
        <Obsolete("Call SetWindowLongPtr instead.", False)>
        <SuppressUnmanagedCodeSecurity>
        <DllImport("User32.dll", EntryPoint:="SetWindowLong", SetLastError:=True)>
        Friend Shared Function SetWindowLong(hWnd As IntPtr,
               <MarshalAs(UnmanagedType.I4)> flags As WindowLongValues,
                                             newLong As UInteger
        ) As UInteger
        End Function

        ''' <summary>
        ''' Changes an attribute of the specified window.
        ''' <para></para>
        ''' The function also sets a value at the specified offset in the extra window memory.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms644898%28v=vs.85%29.aspx"/>
        ''' </remarks>
        '''
        ''' <param name="hWnd">
        ''' A handle to the window and, indirectly, the class to which the window belongs.
        ''' <para></para>
        ''' The <see cref="SetWindowLongPtr"/> function fails if the process that owns the window specified by the 
        ''' <paramref name="hWnd"/> parameter is at a higher process privilege in the 
        ''' <c>UIPI</c> hierarchy than the process the calling thread resides in.
        ''' </param>
        ''' 
        ''' <param name="value">
        ''' The value to be set. 
        ''' </param>
        ''' 
        ''' <param name="newValue">
        ''' The replacement value.
        ''' </param>
        '''
        ''' <returns>
        ''' If the function succeeds, the return value is the previous value of the specified offset.
        ''' <para></para>
        ''' If the function fails, the return value is <see cref="IntPtr.Zero"/>.
        ''' <para></para>
        ''' If the previous value is <see cref="IntPtr.Zero"/> and the function succeeds, the return value is <see cref="IntPtr.Zero"/>.
        ''' <para></para>
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>. 
        ''' </returns>
        <SuppressUnmanagedCodeSecurity>
        <DllImport("User32.dll", EntryPoint:="SetWindowLongPtr", SetLastError:=True)>
        Friend Shared Function SetWindowLongPtr(hWnd As IntPtr,
                  <MarshalAs(UnmanagedType.I4)> value As WindowLongValues,
              <MarshalAs(UnmanagedType.SysInt)> newValue As IntPtr
        ) As <MarshalAs(UnmanagedType.SysInt)> IntPtr
        End Function

        ''' <summary>
        ''' Sends the specified message to a window or windows.
        ''' <para></para>
        ''' The SendMessage function calls the window procedure for the specified window
        ''' and does not return until the window procedure has processed the message.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms644950%28v=vs.85%29.aspx"/>
        ''' </remarks>
        '''
        ''' <param name="hWnd">
        ''' A handle to the window whose window procedure will receive the message.
        ''' </param>
        ''' 
        ''' <param name="msg">
        ''' The message to be sent.
        ''' </param>
        ''' 
        ''' <param name="wParam">
        ''' Additional message-specific information.
        ''' </param>
        ''' 
        ''' <param name="lParam">
        ''' Additional message-specific information.
        ''' </param>
        '''
        ''' <returns>
        ''' The return value specifies the result of the message processing; it depends on the message sent.
        ''' </returns>
        <DllImport("user32.dll", SetLastError:=True)>
        Friend Shared Function SendMessage(hWnd As IntPtr,
                                    msg As Integer,
                                    wParam As IntPtr,
                                    lParam As IntPtr
        ) As IntPtr
        End Function

        ''' <summary>
        ''' Fills a rectangle by using the specified brush. 
        ''' <para></para>
        ''' This function includes the left and top borders, but excludes the right and bottom borders of the rectangle.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/winuser/nf-winuser-fillrect"/>
        ''' </remarks>
        '''
        ''' <param name="hDC">
        ''' A handle to the device context.
        ''' </param>
        ''' 
        ''' <param name="refRect">
        ''' A <see cref="Rectangle"/> that contains the logical coordinates of the rectangle to be filled.
        ''' </param>
        ''' 
        ''' <param name="hBrush">
        ''' A handle to the brush used to fill the rectangle.
        ''' </param>
        '''
        ''' <returns>
        ''' If the function succeeds, the return value is <see langword="True"/>.
        ''' <para></para>
        ''' If the function fails, the return value is <see langword="False"/>. 
        ''' <para></para>
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>.
        ''' </returns>
        <DllImport("user32.dll", SetLastError:=True)>
        Friend Shared Function FillRect(hDC As IntPtr,
                                        ByRef refRect As Rectangle,
                                        hBrush As IntPtr
        ) As Boolean
        End Function

        ''' <summary>
        ''' Adds a rectangle to the specified window's update region. 
        ''' <para></para>
        ''' The update region represents the portion of the window's client area that must be redrawn.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/winuser/nf-winuser-invalidaterect"/>
        ''' </remarks>
        '''
        ''' <param name="hWnd">
        ''' A handle to the window whose update region has changed. 
        ''' <para></para>
        ''' If this parameter is <see cref="IntPtr.Zero"/>, the system invalidates and redraws all windows, 
        ''' not just the windows for this application, and sends the 
        ''' <see cref="WindowMessages.WM_EraseBkgnd"/> and <see cref="WindowMessages.WM_NcPaint"/> messages before the function returns. 
        ''' <para></para>
        ''' Setting this parameter to <see cref="IntPtr.Zero"/> is not recommended.
        ''' </param>
        ''' 
        ''' <param name="refRect">
        ''' A <see cref="NativeRectangle"/> structure that contains the client coordinates of the 
        ''' rectangle to be added to the update region. 
        ''' <para></para>
        ''' If this parameter is <see langword="Nothing"/>, the entire client area is added to the update region.
        ''' </param>
        ''' 
        ''' <param name="erase">
        ''' Specifies whether the background within the update region is to be erased when the update region is processed. 
        ''' <para></para>
        ''' If this parameter is <see langword="True"/>, 
        ''' the background is erased when the <see cref="NativeMethods.BeginPaint"/> function is called. 
        ''' If this parameter is <see langword="False"/>, the background remains unchanged.
        ''' </param>
        '''
        ''' <returns>
        ''' If the function succeeds, the return value is <see langword="True"/>.
        ''' If the function fails, the return value is <see langword="False"/>. 
        ''' <para></para>
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>.
        ''' </returns>
        <DllImport("user32.dll", SetLastError:=True)>
        Friend Shared Function InvalidateRect(hWnd As IntPtr,
                                              ByRef refRect As NativeRectangle,
              <MarshalAs(UnmanagedType.Bool)> [erase] As Boolean
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Updates the client area of the specified window 
        ''' by sending a <see cref="WindowMessages.WM_Paint"/> message to the window if the window's update region is not empty. 
        ''' <para></para>
        ''' The function sends a <see cref="WindowMessages.WM_Paint"/> message directly to the window procedure of the specified window, 
        ''' bypassing the application queue. If the update region is empty, no message is sent.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/winuser/nf-winuser-updatewindow"/>
        ''' </remarks>
        '''
        ''' <param name="hWnd">
        ''' A handle to the window to be updated.
        ''' </param>
        '''
        ''' <returns>
        ''' If the function succeeds, the return value is <see langword="True"/>.
        ''' If the function fails, the return value is <see langword="False"/>. 
        ''' <para></para>
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>.
        ''' </returns>
        <DllImport("user32.dll", SetLastError:=True)>
        Friend Shared Function UpdateWindow(hWnd As IntPtr
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Retrieves a handle to a device context (DC) for the client area of a 
        ''' specified window or for the entire screen. 
        ''' <para></para>
        ''' You can use the returned handle in subsequent GDI functions to draw in the DC. 
        ''' <para></para>
        ''' The device context is an opaque data structure, whose values are used internally by GDI.
        ''' <para></para>
        ''' This function is an extension to the <see cref="GetDC"/> function, 
        ''' which gives an application more control over how and 
        ''' whether clipping occurs in the client area.
        ''' </summary>
        ''' 
        ''' <remarks>
        ''' <see href="https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getdcex"/>
        ''' </remarks>
        '''
        ''' <param name="hWnd">
        ''' A handle to the window whose DC is to be retrieved. 
        ''' <para></para>
        ''' If this value is NULL, <see cref="GetDCEx"/> retrieves the DC for the entire screen.
        ''' </param>
        ''' 
        ''' <param name="hRgnClip">
        ''' A clipping region that may be combined with the visible region of the DC.
        ''' <para></para>
        ''' If the value of <paramref name="flags"/> prameter is <see cref="GetDCExFlags.IntersectRegion"/> 
        ''' or <see cref="GetDCExFlags.ExcludeRegion"/>, then the operating system assumes ownership of the 
        ''' region and will automatically delete it when it is no longer needed. 
        ''' In this case, the application should not use or delete the region after a successful call to <see cref="GetDCEx"/>.
        ''' </param>
        ''' 
        ''' <param name="flags">
        ''' Specifies how the DC is created. 
        ''' </param>
        ''' 
        ''' <returns>
        ''' If the function succeeds, the return value is the handle to the DC for the specified window.
        ''' <para></para>
        ''' If the function fails, the return value is NULL. 
        ''' <para></para>
        ''' An invalid value for the <paramref name="hWnd"/> parameter will cause the function to fail.
        ''' </returns>
        <DllImport("user32.dll", SetLastError:=False, ExactSpelling:=True)>
        Friend Shared Function GetDCEx(<[In], [Optional]> hWnd As IntPtr,
                                <[In], [Optional]> hRgnClip As IntPtr,
                                                   flags As GetDCExFlags
        ) As IntPtr
        End Function

#End Region

    End Class

End Namespace

#End Region
