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

Imports DevCase.Interop.Unmanaged.Win32.Structures

#End Region

#Region " SHGetFileInfoFlags "

Namespace DevCase.Interop.Unmanaged.Win32.Enums

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Flags to use with the <see cref="NativeMethods.SHGetFileInfo"/> function.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' <see href="https://msdn.microsoft.com/es-es/library/windows/desktop/bb762179(v=vs.85).aspx"/>
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    <Flags>
    Friend Enum SHGetFileInfoFlags As Integer

        ''' <summary>
        ''' Retrieve the handle to the icon that represents the file and the index of the 
        ''' icon within the system image list. 
        ''' <para></para>
        ''' The handle is copied to the <see cref="ShellFileInfo.IconHandle"/> member of the 
        ''' structure specified by <c>refShellFileInfo</c> parameter, 
        ''' and the index is copied to the <see cref="ShellFileInfo.IconIndex"/> member.
        ''' </summary>
        Icon = &H100

        ''' <summary>
        ''' Retrieve the display name for the file, which is the name as it appears in Windows Explorer. 
        ''' <para></para>
        ''' The name is copied to the <see cref="ShellFileInfo.DisplayName"/> member of the 
        ''' structure specified in <c>refShellFileInfo</c> parameter. 
        ''' <para></para>
        ''' The returned display name uses the long file name, if there is one, 
        ''' rather than the 8.3 form of the file name. 
        ''' <para></para>
        ''' Note that the display name can be affected by settings such as whether extensions are shown.
        ''' </summary>
        DisplayName = &H200

        ''' <summary>
        ''' Retrieve the string that describes the file's type. 
        ''' <para></para>
        ''' The string is copied to the <see cref="ShellFileInfo.TypeName"/> member of the 
        ''' structure specified in <c>refShellFileInfo</c> parameter.
        ''' </summary>
        TypeName = &H400

        ''' <summary>
        ''' Retrieve the item attributes. 
        ''' <para></para>
        ''' The attributes are copied to the <see cref="ShellFileInfo.Attributes"/> member of the 
        ''' structure specified in the <c>refShellFileInfo</c> parameter. 
        ''' <para></para>
        ''' These are the same attributes that are obtained from <c>IShellFolder::GetAttributesOf</c>.
        ''' </summary>
        Attributes = &H800

        ''' <summary>
        ''' Retrieve the name of the file that contains the icon representing the file specified by <c>path</c> parameter, 
        ''' as returned by the <c>IExtractIcon::GetIconLocation</c> method of the file's icon handler. 
        ''' <para></para>
        ''' Also retrieve the icon index within that file. 
        ''' <para></para>
        ''' The name of the file containing the icon is copied to the <see cref="ShellFileInfo.DisplayName"/> member 
        ''' of the structure specified by <c>refShellFileInfo</c> parameter. 
        ''' <para></para>
        ''' The icon's index is copied to that structure's <see cref="ShellFileInfo.IconIndex"/> member.
        ''' </summary>
        IconLocation = &H1000

        ''' <summary>
        ''' Retrieve the type of the executable file if <c>path</c> parameter identifies an executable file. 
        ''' <para></para>
        ''' The information is packed into the return value. 
        ''' <para></para>
        ''' This flag cannot be specified with any other flags.
        ''' </summary>
        ExeType = &H2000

        ''' <summary>
        ''' Retrieve the index of a system image list icon. 
        ''' <para></para>
        ''' If successful, the index is copied to the iIcon member of <c>refShellFileInfo</c>. 
        ''' <para></para>
        ''' The return value is a handle to the system image list. 
        ''' <para></para>
        ''' Only those images whose indices are successfully copied to 
        ''' <see cref="ShellFileInfo.IconIndex"/> are valid. 
        ''' <para></para>
        ''' Attempting to access other images in the system image list will result in undefined behavior
        ''' </summary>
        SysIconIndex = &H4000

        ''' <summary>
        ''' Modify <see cref="SHGetFileInfoFlags.Icon"/>, causing the function to add the link overlay to the file's icon. 
        ''' <para></para>
        ''' The <see cref="SHGetFileInfoFlags.Icon"/> flag must also be set.
        ''' </summary>
        LinkOverlay = &H8000

        ''' <summary>
        ''' Modify <see cref="SHGetFileInfoFlags.Icon"/>, causing the function to blend the file's icon with the 
        ''' system highlight color. 
        ''' <para></para>
        ''' The <see cref="SHGetFileInfoFlags.Icon"/> flag must also be set.
        ''' </summary>
        Selected = &H10000

        ''' <summary>
        ''' Modify <see cref="SHGetFileInfoFlags.Attributes"/> to indicate that the 
        ''' <see cref="ShellFileInfo.Attributes"/> member of the structure at <c>refShellFileInfo</c> parameter 
        ''' contains the specific attributes that are desired. 
        ''' <para></para>
        ''' These attributes are passed to <c>IShellFolder::GetAttributesOf</c>. 
        ''' <para></para>
        ''' If this flag is not specified, <c>0xFFFFFFFF</c> is passed to <c>IShellFolder::GetAttributesOf</c>, 
        ''' requesting all attributes. 
        ''' <para></para>
        ''' This flag cannot be specified with the <see cref="SHGetFileInfoFlags.Icon"/> flag.
        ''' </summary>
        AttributeSpecified = &H20000

        ''' <summary>
        ''' Modify <see cref="SHGetFileInfoFlags.Icon"/>, causing the function to retrieve the file's large icon. 
        ''' <para></para>
        ''' The image size is normally 32x32 pixels. 
        ''' However, if the Use large icons option is selected from the Effects section of the Appearance tab in
		''' Display Properties, the image is 48x48 pixels.
        ''' <para></para>
        ''' The <see cref="SHGetFileInfoFlags.Icon"/> flag must also be set.
        ''' </summary>
        IconSizeLarge = &H0

        ''' <summary>
        ''' Modify <see cref="SHGetFileInfoFlags.Icon"/>, causing the function to retrieve the file's small icon. 
        ''' <para></para>
        ''' Also used to modify <see cref="SHGetFileInfoFlags.SysIconIndex"/>, causing the function to return the
        ''' handle to the system image list that contains small icon images. 
        ''' <para></para>
        ''' The image is the shell standard small icon size of 16x16, but the size can be customized by the user.
        ''' <para></para>
        ''' The <see cref="SHGetFileInfoFlags.Icon"/> and/or <see cref="SHGetFileInfoFlags.SysIconIndex"/> flag must also be set.
        ''' </summary>
        IconSizeSmall = &H1

        ''' <summary>
        ''' Modify <see cref="SHGetFileInfoFlags.Icon"/>, causing the function to retrieve the file's open icon. 
        ''' <para></para>
        ''' Also used to modify <see cref="SHGetFileInfoFlags.SysIconIndex"/>, causing the function to return the 
        ''' handle to the system image list that contains the file's small open icon. 
        ''' <para></para>
        ''' A container object displays an open icon to indicate that the container is open. 
        ''' <para></para>
        ''' The image is the Shell standard extra-large icon size. 
        ''' This is typically 48x48, but the size can be customized by the user.
        ''' <para></para>
        ''' The <see cref="SHGetFileInfoFlags.Icon"/> and/or <see cref="SHGetFileInfoFlags.SysIconIndex"/> flag must also be set.
        ''' </summary>
        IconSizeExtraLarge = &H2

        ''' <summary>
        ''' Modify <see cref="SHGetFileInfoFlags.Icon"/>, causing the function to retrieve a Shell-sized icon. 
        ''' <para></para>
        ''' If this flag is not specified the function sizes the icon according to the system metric values. 
        ''' <para></para>
        ''' The image is normally 256x256 pixels.
        ''' <para></para>
        ''' The <see cref="SHGetFileInfoFlags.Icon"/> flag must also be set.
        ''' </summary>
        IconSizeJumbo = &H4

        ''' <summary>
        ''' Indicate that <c>path</c> parameter is the address of an <c>ITEMIDLIST</c> structure 
        ''' rather than a path name.
        ''' </summary>
        PIDL = &H8

        ''' <summary>
        ''' Indicates that the function should not attempt to access the file specified by <c>path</c> parameter. 
        ''' <para></para>
        ''' Rather, it should act as if the file specified by <c>path</c> parameter exists with the 
        ''' file attributes passed in <c>fileAttributes</c> parameter. 
        ''' <para></para>
        ''' This flag cannot be combined with the <see cref="SHGetFileInfoFlags.Attributes"/>, 
        ''' <see cref="SHGetFileInfoFlags.ExeType"/>, or <see cref="SHGetFileInfoFlags.PIDL"/> flags.
        ''' </summary>
        UseFileAttributes = &H10

        ''' <summary>
        ''' Apply the appropriate overlays to the file's icon. 
        ''' <para></para>
        ''' The <see cref="SHGetFileInfoFlags.Icon"/> flag must also be set.
        ''' </summary>
        AddOverlays = &H20

        ''' <summary>
        ''' Return the index of the overlay icon. 
        ''' <para></para>
        ''' The value of the overlay index is returned in the upper eight bits of the 
        ''' <see cref="ShellFileInfo.IconIndex"/> member of the structure specified by <c>refShellFileInfo</c> parameter. 
        ''' <para></para>
        ''' This flag requires that the <see cref="SHGetFileInfoFlags.Icon"/> be set as well.
        ''' </summary>
        OverlayIndex = &H40

    End Enum

End Namespace

#End Region
