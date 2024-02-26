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

#Region " ShellItemGetDisplayName "

Namespace DevCase.Interop.Unmanaged.Win32.Enums

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Requests the form of an item's display name to retrieve 
    ''' through IShellItem.GetDisplayName and <see cref="NativeMethods.SHGetNameFromIDList"/> functions.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/shobjidl_core/ne-shobjidl_core-_sigdn"/>
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    Friend Enum ShellItemGetDisplayName As UInteger ' SIGDN

        ''' <summary>
        ''' Returns the editing name relative to the desktop. In UI this name is suitable for display to the user.
        ''' </summary>
        DesktopAbsoluteEditing = &H8004C000UI

        ''' <summary>
        ''' Returns the parsing name relative to the desktop. This name is not suitable for use in UI.
        ''' </summary>
        DesktopAbsoluteParsing = &H80028000UI

        ''' <summary>
        ''' Returns the item's file system path, if it has one. 
        ''' <para></para>
        ''' Only items that report SFGAO_FILESYSTEM have a file system path. 
        ''' <para></para>
        ''' When an item does not have a file system path, a call to IShellItem::GetDisplayName on that item will fail. 
        ''' <para></para>
        ''' In UI this name is suitable for display to the user in some cases, but note that it might not be specified for all items.
        ''' </summary>
        FileSystemPath = &H80058000UI

        ''' <summary>
        ''' Returns the display name relative to the parent folder. In UI this name is generally ideal for display to the user.
        ''' </summary>
        NormalDisplay = 0

        ''' <summary>
        ''' Returns the path relative to the parent folder.
        ''' </summary>
        ParentRelative = &H80080001UI

        ''' <summary>
        ''' Returns the editing name relative to the parent folder. In UI this name is suitable for display to the user.
        ''' </summary>
        ParentRelativeEditing = &H80031001UI

        ''' <summary>
        ''' Returns the path relative to the parent folder in a friendly format as displayed in an address bar. 
        ''' <para></para>
        ''' This name is suitable for display to the user.
        ''' </summary>
        ParentRelativeForAddressBar = &H8007C001UI

        ''' <summary>
        ''' Not documented. Introduced in Windows 8.
        ''' </summary>
        ParentRelativeForUI = &H80094001UI

        ''' <summary>
        ''' Returns the parsing name relative to the parent folder. This name is not suitable for use in UI.
        ''' </summary>
        ParentRelativeParsing = &H80018001UI

        ''' <summary>
        ''' Returns the item's URL, if it has one. 
        ''' <para></para>
        ''' Some items do not have a URL, and in those cases a call to IShellItem.GetDisplayName will fail. 
        ''' <para></para>
        ''' This name is suitable for display to the user in some cases, but note that it might not be specified for all items.
        ''' </summary>
        Url = &H80068000UI

    End Enum

End Namespace

#End Region
