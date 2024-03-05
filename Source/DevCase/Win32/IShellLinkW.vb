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

Imports System.Runtime.InteropServices
Imports System.Text

Imports DevCase.Interop.Unmanaged.Win32.Enums
Imports DevCase.Interop.Unmanaged.Win32.Structures

#End Region

#Region " IShellLink (Unicode) "

Namespace DevCase.Interop.Unmanaged.Win32.Interfaces

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' The <c>IShellLink</c> interface allows Shell links to be created, modified, or resolved.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' <see href="https://msdn.microsoft.com/es-es/library/windows/desktop/bb774950%28v=vs.85%29.aspx"/>
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    <ComImport>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <Guid("000214F9-0000-0000-C000-000000000046")>
    Friend Interface IShellLinkW

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves the path and file name of a Shell link object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="file">
        ''' The address of a buffer that receives the path and file name of the target of the Shell link object.
        ''' </param>
        ''' 
        ''' <param name="maxPath">
        ''' The size, in characters, of the buffer pointed to by the pszFile parameter, 
        ''' including the terminating null character. 
        ''' <para></para>
        ''' The maximum path size that can be returned is <c>MAX_PATH</c>.
        ''' </param>
        ''' 
        ''' <param name="refWin32FindData">
        ''' A pointer to a WIN32_FIND_DATA structure that receives information about the target of the Shell link object.
        ''' <para></para>
        ''' If this parameter is NULL, then no additional information is returned.
        ''' </param>
        ''' 
        ''' <param name="flags">
        ''' Flags that specify the type of path information to retrieve.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Sub GetPath(<Out, MarshalAs(UnmanagedType.LPWStr)> file As StringBuilder,
                                                           maxPath As Integer,
                                                     ByRef refWin32FindData As Win32FindDataW,
                                                           flags As IShellLinkGetPathFlags)

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves the list of item identifiers for a Shell link object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="refPidl">
        ''' When this method returns, contains the address of a PIDL.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Sub GetIDList(ByRef refPidl As IntPtr)

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sets the pointer to an item identifier list (PIDL) for a Shell link object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidl">
        ''' The object's fully qualified PIDL.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Sub SetIDList(pidl As IntPtr)

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves the description string for a Shell link object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="name">
        ''' A pointer to the buffer that receives the description string.
        ''' </param>
        ''' 
        ''' <param name="maxName">
        ''' The maximum number of characters to copy to the buffer pointed to by the pszName parameter.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Sub GetDescription(<Out, MarshalAs(UnmanagedType.LPWStr)> name As StringBuilder, maxName As Integer)

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sets the description for a Shell link object.
        ''' The description can be any application-defined string.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="name">
        ''' A pointer to a buffer containing the new description string.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Sub SetDescription(<MarshalAs(UnmanagedType.LPWStr)> name As String)

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves the name of the working directory for a Shell link object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="dir">
        ''' The address of a buffer that receives the name of the working directory.
        ''' </param>
        ''' 
        ''' <param name="maxPath">
        ''' The maximum number of characters to copy to the buffer pointed to by the pszDir parameter.
        ''' <para></para>
        ''' The name of the working directory is truncated if it is longer than the maximum specified by this parameter.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Sub GetWorkingDirectory(<Out, MarshalAs(UnmanagedType.LPWStr)> dir As StringBuilder,
                                                                       maxPath As Integer)

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sets the name of the working directory for a Shell link object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="dir">
        ''' The address of a buffer that contains the name of the new working directory.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Sub SetWorkingDirectory(<MarshalAs(UnmanagedType.LPWStr)> dir As String)

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves the command-line arguments associated with a Shell link object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="args">
        ''' A pointer to the buffer that, when this method returns successfully, receives the command-line arguments.
        ''' </param>
        ''' 
        ''' <param name="maxPath">
        ''' The maximum number of characters that can be copied to the buffer supplied by the pszArgs parameter.
        ''' <para></para>
        ''' In the case of a Unicode string, there is no limitation on maximum string length.
        ''' <para></para>
        ''' In the case of an ANSI string, the maximum length of the returned string varies depending on the 
        ''' version of Windows.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Sub GetArguments(<Out, MarshalAs(UnmanagedType.LPWStr)> args As StringBuilder,
                                                                maxPath As Integer)

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sets the command-line arguments for a Shell link object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="args">
        ''' A pointer to a buffer that contains the new command-line arguments.
        ''' <para></para>
        ''' In the case of a Unicode string, there is no limitation on maximum string length.
        ''' <para></para>
        ''' In the case of an ANSI string, the maximum length of the returned string varies depending on the 
        ''' version of Windows.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Sub SetArguments(<MarshalAs(UnmanagedType.LPWStr)> args As String)

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves the hot key for a Shell link object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="refHotkey">
        ''' The address of the keyboard shortcut.
        ''' <para></para>
        ''' The virtual key code is in the low-order byte, and the hotkey modifier flags are in the high-order byte.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Sub GetHotkey(ByRef refHotkey As UShort)

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sets a hot key for a Shell link object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hotkey">
        ''' The new keyboard shortcut.
        ''' <para></para>
        ''' The virtual key code is in the low-order byte, and the modifier flags are in the high-order byte.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Sub SetHotkey(hotkey As UShort)

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves the ShowWindowFlags for a Shell link object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="refWindowState">
        ''' A <see cref="NativeWindowState"/> Flags.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Sub GetShowCmd(ByRef refWindowState As NativeWindowState)

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sets the show command for a Shell link object.
        ''' <para></para>
        ''' The show command sets the initial show state of the window.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="windowState">
        ''' A <see cref="NativeWindowState"/> flags.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Sub SetShowCmd(windowState As NativeWindowState)

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves the location (path and index) of the icon for a Shell link object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="iconPath">
        ''' The address of a buffer that receives the path of the file containing the icon.
        ''' </param>
        ''' 
        ''' <param name="iconPathSize">
        ''' The maximum number of characters to copy to the buffer pointed to by the <paramref name="iconPath"/> parameter.
        ''' </param>
        ''' 
        ''' <param name="refIconIndex">
        ''' The address of a value that receives the index of the icon.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Sub GetIconLocation(<Out, MarshalAs(UnmanagedType.LPWStr)> iconPath As StringBuilder,
                                                                   iconPathSize As Integer,
                                                             ByRef refIconIndex As Integer)

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sets the location (path and index) of the icon for a Shell link object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="iconPath">
        ''' The address of a buffer to contain the path of the file containing the icon.
        ''' </param>
        ''' 
        ''' <param name="iconIndex">
        ''' The index of the icon.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Sub SetIconLocation(<MarshalAs(UnmanagedType.LPWStr)> iconPath As String,
                                                              iconIndex As Integer)

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sets the relative path to the Shell link object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pathRelative">
        ''' The address of a buffer that contains the fully-qualified path of the shortcut file, 
        ''' relative to which the shortcut resolution should be performed. 
        ''' <para></para>
        ''' It should be a file name, not a folder name.
        ''' </param>
        ''' 
        ''' <param name="reserved">
        ''' Reserved.
        ''' <para></para>
        ''' Set this parameter to zero.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Sub SetRelativePath(<MarshalAs(UnmanagedType.LPWStr)> pathRelative As String,
                                                              reserved As Integer)

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Attempts to find the target of a Shell link, even if it has been moved or renamed.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hWnd">
        ''' A handle to the window that the Shell will use as the parent for a dialog box.
        ''' <para></para>
        ''' The Shell displays the dialog box if it needs to prompt the user for more information while resolving a link.
        ''' </param>
        ''' 
        ''' <param name="flags">
        ''' Action flags.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Sub Resolve(hWnd As IntPtr,
                    flags As IShellLinkResolveFlags)

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sets the path and file name of a Shell link object
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="file">
        ''' The address of a buffer that contains the new path.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Sub SetPath(file As String)

    End Interface

End Namespace

#End Region
