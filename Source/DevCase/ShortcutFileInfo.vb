﻿' This source-code is freely distributed as part of "DevCase Class Library .NET Developers".
' 
' Maybe you would like to consider to buy this powerful set of libraries to support me. 
' You can do loads of things with my apis for a big amount of diverse thematics.
' 
' Here is a link to the purchase page:
' https://codecanyon.net/item/elektrokit-class-library-for-net/19260282
' 
' Thank you.


#Region " Public Members Summary "

#Region " Constructors "

' ShortcutInfo.New(String)
' ShortcutInfo.New(FileInfo)

#End Region

#Region " Properties "

' Attributes As FileAttributes
' CreationTime As Date
' CreationTimeUtc As Date
' Description As String
' Directory As DirectoryInfo
' DirectoryName As String
' Exists As Boolean
' Extension As String
' FullName As String
' Hotkey As Keys
' Icon As String
' IconIndex As Integer
' IsReadOnly As Boolean
' LastAccessTime As Date
' LastAccessTimeUtc As Date
' LastWriteTime As Date
' LastWriteTimeUtc As Date
' Length As Long
' Name As String
' Target As String
' TargetArguments As String
' TargetDisplayName As String
' ViewMode As Boolean
' WindowState As ShortcutWindowState
' WorkingDirectory As String

#End Region

#Region " Methods "

' CopyTo(String, Boolean) As ShortcutFileInfo
' Create()
' Decrypt()
' Delete()
' Encrypt()
' GetAccessControl() As FileSecurity
' GetAccessControl(AccessControlSections) As FileSecurity
' MoveTo(String)
' Open(FileMode) As FileStream
' Open(FileMode, FileAccess) As FileStream
' Open(FileMode, FileAccess, FileShare) As FileStream
' OpenRead() As FileStream
' OpenWrite() As FileStream
' Refresh()
' Resolve(IntPtr, IShellLinkResolveFlags)
' ToString() As String

#End Region

#End Region

#Region " Usage Examples "

' Dim lnk As New ShortcutFileInfo("C:\Test Shortcut.lnk")
' lnk.Create()
' 
' With lnk
'     .Attributes = FileAttributes.Normal
'     .Description = "My shortcut description."
'     .Hotkey = Keys.Shift Or Keys.Alt Or Keys.Control Or Keys.F1
'     .Icon = "Shell32.dll"
'     .IconIndex = 0
'     .Target = "C:\Windows\Notepad.exe"
'     .TargetArguments = """C:\Windows\win.ini"""
'     .WindowState = ShortcutWindowState.Normal
'     .WorkingDirectory = Path.GetDirectoryName(lnk.Target)
' End With
' 
' lnk.ViewMode = True
' Me.PropertyGrid1.SelectedObject = lnk

#End Region

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.CodeDom
Imports System.ComponentModel
Imports System.Drawing.Design
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security
Imports System.Security.AccessControl
Imports System.Security.Policy
Imports System.Text
Imports System.Windows.Forms.Design
Imports System.Xml.Serialization

Imports DevCase.Core.Design
Imports DevCase.Interop.Unmanaged.Win32
Imports DevCase.Interop.Unmanaged.Win32.Enums
Imports DevCase.Interop.Unmanaged.Win32.Interfaces

Imports Newtonsoft.Json.Linq



#End Region
#Region " ShortcutFileInfo "

Namespace DevCase.Core.IO

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Provides information about a shortcut (.lnk) file.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim lnk As New ShortcutFileInfo("C:\Test Shortcut.lnk")
    ''' lnk.Create()
    ''' 
    ''' With lnk
    '''     .Attributes = FileAttributes.Normal
    '''     .Description = "My shortcut description."
    '''     .Hotkey = Keys.Shift Or Keys.Alt Or Keys.Control Or Keys.F1
    '''     .Icon = "Shell32.dll"
    '''     .IconIndex = 0
    '''     .Target = "C:\Windows\Notepad.exe"
    '''     .TargetArguments = """C:\Windows\win.ini"""
    '''     .WindowState = ShortcutWindowState.Normal
    '''     .WorkingDirectory = Path.GetDirectoryName(lnk.Target)
    ''' End With
    ''' 
    ''' lnk.ViewMode = True
    ''' Me.PropertyGrid1.SelectedObject = lnk
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    <Serializable>
    <XmlRoot(NameOf(ShortcutFileInfo))>
    <TypeConverter(GetType(ExpandableObjectConverter))>
    <Category(NameOf(ShortcutFileInfo))>
    <Description("Provides information about a shortcut (.lnk) file.")>
    <DefaultProperty(NameOf(ShortcutFileInfo.Target))>
    Public NotInheritable Class ShortcutFileInfo : Inherits FileSystemInfo

#Region " Private Fields "

        Private Const maxArgumentsLength As Integer = 1023
        Private Const maxDescriptionLength As Integer = 259
        Private Const maxIconLength As Integer = 259
        Private Const maxTargetLength As Integer = 259
        Private Const maxWorkingDirLength As Integer = 259

#End Region

#Region " Properties "

#Region " File Info "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the file name of the link file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The name of the link file.")>
        <DisplayName("File Name")>
        <Category("File Info")>
        Public Overrides ReadOnly Property Name As String
            Get
                Return Path.GetFileName(MyBase.FullName)
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the full path of the link file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The full path of the link file.")>
        <DisplayName("Full Path")>
        <Category("File Info")>
        Public Overrides ReadOnly Property FullName As String
            Get
                Return MyBase.FullName
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the file size, in bytes, of the current link file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The file attributes.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The file size of the link file.")>
        <DisplayName("Length")>
        <Category("File Info")>
        <Browsable(True)>
        <TypeConverter(GetType(FileSizeConverter))>
        Public ReadOnly Property Length As Long
            Get
                Return New FileInfo(Me.FullName).Length
            End Get
        End Property

#End Region

#Region " File Info (non-browsable)"

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the file attributes of the link file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The file attributes.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The file attributes of the link file.")>
        <DisplayName("Attributes")>
        <Category("File Info")>
        <Browsable(False)>
        Public Shadows Property Attributes As FileAttributes
            Get
                Return MyBase.Attributes
            End Get
            Set(value As FileAttributes)
                If (value <> MyBase.Attributes) Then
                    MyBase.Attributes = value
                End If
            End Set
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines if the link file is read-only.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if the link file is read-only; otherwise, <see langword="False"/>.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("A value that determines if the link file is read-only.")>
        <DisplayName("Is Read-Only")>
        <Category("File Info")>
        <Browsable(False)>
        Public Property IsReadOnly As Boolean
            Get
                Return Me.Attributes.HasFlag(FileAttributes.ReadOnly)
            End Get
            Set(value As Boolean)
                Me.Attributes = (Me.Attributes Or FileAttributes.ReadOnly)
            End Set
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a value indicating whether the link file exists.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if the link file exists; otherwise, <see langword="False"/>.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("A value indicating whether the link file exists.")>
        <DisplayName("Exists")>
        <Category("File Info")>
        <Browsable(False)>
        Public Overrides ReadOnly Property Exists As Boolean
            Get
                Return File.Exists(Me.FullName)
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets an instance of the parent directory of the link file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The parent directory of the link file.")>
        <DisplayName("Directory")>
        <Category("File Info")>
        <Browsable(False)>
        Public ReadOnly Property Directory As DirectoryInfo
            Get
                Return New FileInfo(Me.FullName).Directory
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a string representing the shortcut directory's full path.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("A string representing the shortcut directory's full path.")>
        <DisplayName("Directory Name")>
        <Category("File Info")>
        <Browsable(False)>
        Public ReadOnly Property DirectoryName As String
            Get
                Return New FileInfo(Me.FullName).DirectoryName
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a string representing the extension part of the link file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The extension.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("A string representing the extension part of the link file.")>
        <DisplayName("Extension")>
        <Category("File Info")>
        <Browsable(False)>
        Public Shadows ReadOnly Property Extension As String
            Get
                Return MyBase.Extension
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the creation time, in coordinated universal time (UTC) of the link file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The creation time UTC.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The creation time, in coordinated universal time (UTC) of the link file.")>
        <DisplayName("Creation Time (UTC)")>
        <Category("File Info")>
        <Browsable(False)>
        Public Shadows Property CreationTimeUtc As Date
            Get
                Return MyBase.CreationTimeUtc
            End Get
            Set(value As Date)
                If (value <> MyBase.CreationTimeUtc) Then
                    MyBase.CreationTimeUtc = value
                End If
            End Set
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the time, in coordinated universal time (UTC), the link file was last accessed.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The last access time UTC.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The time, in coordinated universal time (UTC), the link file was last accessed.")>
        <DisplayName("Last Access Time (UTC)")>
        <Category("File Info")>
        <Browsable(False)>
        Public Shadows Property LastAccessTimeUtc As Date
            Get
                Return MyBase.LastAccessTimeUtc
            End Get
            Set(value As Date)
                If (value <> MyBase.LastAccessTimeUtc) Then
                    MyBase.LastAccessTimeUtc = value
                End If
            End Set
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the time, in coordinated universal time (UTC), the link file was last written to.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The last write time UTC.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The time, in coordinated universal time (UTC), the link file was last written to.")>
        <DisplayName("Last Write Time (UTC)")>
        <Category("File Info")>
        <Browsable(False)>
        Public Shadows Property LastWriteTimeUtc As Date
            Get
                Return MyBase.LastWriteTimeUtc
            End Get
            Set(value As Date)
                If (value <> MyBase.LastWriteTimeUtc) Then
                    MyBase.LastWriteTimeUtc = value
                End If
            End Set
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the creation time of the link file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The creation time.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The creation time of the link file.")>
        <DisplayName("Creation Time")>
        <Category("File Info")>
        <Browsable(False)>
        Public Shadows Property CreationTime As Date
            Get
                Return MyBase.CreationTime
            End Get
            Set(value As Date)
                If (value <> MyBase.CreationTime) Then
                    MyBase.CreationTime = value
                End If
            End Set
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the time the link file was last accessed.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The last access time.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The time the link file was last accessed.")>
        <DisplayName("Last Access Time")>
        <Category("File Info")>
        <Browsable(False)>
        Public Shadows Property LastAccessTime As Date
            Get
                Return MyBase.LastAccessTime
            End Get
            Set(value As Date)
                If (value <> MyBase.LastAccessTime) Then
                    MyBase.LastAccessTime = value
                End If
            End Set
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the time the link file was last written to.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The last write time.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The time the link file was last written to.")>
        <DisplayName("Last Write Time")>
        <Category("File Info")>
        <Browsable(False)>
        Public Shadows Property LastWriteTime As Date
            Get
                Return MyBase.LastWriteTime
            End Get
            Set(value As Date)
                If (value <> MyBase.LastWriteTime) Then
                    MyBase.LastWriteTime = value
                End If
            End Set
        End Property

#End Region

#Region " Shortcut "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the shortcut description.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The shortcut description. Max Length is 259 characters.")>
        <DisplayName("Description")>
        <Category("Shortcut")>
        <DefaultValue("")>
        Public Property Description As String
            Get
                Return Me.description_
            End Get
            Set(value As String)
                If (value <> Me.description_) Then
                    If value.Length > maxDescriptionLength Then
                        value = value.Substring(0, maxDescriptionLength)
                    End If
                    Me.description_ = value
                    Me.WriteLink()
                End If
            End Set
        End Property
        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' ( Backing Field of <see cref="ShortcutFileInfo.Description"/> property. )
        ''' <para></para>
        ''' The shortcut description.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private description_ As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the shortcut hotkey.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The shortcut hotkey.")>
        <DisplayName("Hotkey")>
        <Category("Shortcut")>
        <DefaultValue(Keys.None)>
        <Editor(GetType(ShortcutKeysEditor), GetType(UITypeEditor))>
        Public Property Hotkey As Keys
            Get
                Return Me.hotkey_
            End Get
            Set(value As Keys)
                If (value <> Me.hotkey_) Then
                    Me.hotkey_ = value
                    Me.WriteLink()
                End If
            End Set
        End Property
        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' ( Backing Field of <see cref="ShortcutFileInfo.Hotkey"/> property. )
        ''' <para></para>
        ''' The shortcut hotkey.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private hotkey_ As Keys

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the full path of the icon file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The full path of the icon file. Max Length is 259 characters.")>
        <DisplayName("Icon")>
        <Category("Shortcut")>
        <DefaultValue("")>
        <Editor(GetType(IconFileNameEditor), GetType(UITypeEditor))>
        Public Property Icon As String
            Get
                Return Me.icon_
            End Get
            Set(value As String)
                If (value <> Me.icon_) Then
                    If value.Length > maxIconLength Then
                        value = value.Substring(0, maxIconLength)
                    End If
                    Me.icon_ = value
                    Me.WriteLink()
                End If
            End Set
        End Property
        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' ( Backing Field of <see cref="ShortcutFileInfo.Icon"/> property. )
        ''' <para></para>
        ''' The full path of the icon file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private icon_ As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the image index within the icon file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Editor(GetType(ShortcutFileInfoIconIndexEditor), GetType(UITypeEditor))>
        <Description("The image index within the icon file.")>
        <DisplayName("Icon Index")>
        <Category("Shortcut")>
        <DefaultValue(0)>
        Public Property IconIndex As Integer
            Get
                Return Me.iconIndex_
            End Get
            Set(value As Integer)
                If (value <> Me.iconIndex_) Then
                    Me.iconIndex_ = value
                    Me.WriteLink()
                End If
            End Set
        End Property
        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' ( Backing Field of <see cref="ShortcutFileInfo.IconIndex"/> property. )
        ''' <para></para>
        ''' The image index within the icon file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private iconIndex_ As Integer

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the window state for the target file or directory.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The window state for the target file or directory.")>
        <DisplayName("Window State")>
        <Category("Shortcut")>
        <DefaultValue(ShortcutWindowState.Normal)>
        Public Property WindowState As ShortcutWindowState
            Get
                Return Me.windowState_
            End Get
            Set(value As ShortcutWindowState)
                If (value <> Me.windowState_) Then
                    Me.windowState_ = value
                    Me.WriteLink()
                End If
            End Set
        End Property
        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' ( Backing Field of <see cref="ShortcutFileInfo.WindowState"/> property. )
        ''' <para></para>
        ''' The window state of the target file or directory.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private windowState_ As ShortcutWindowState

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the full path of the target file or directory.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The full path of the target file or directory. Max Length is 259 characters.")>
        <DisplayName("Target")>
        <Category("Shortcut")>
        <DefaultValue("")>
        <Editor(GetType(FileOrFolderNameEditor), GetType(UITypeEditor))>
        Public Property Target As String
            Get
                Return Me.target_
            End Get
            Set(value As String)
                If (value <> Me.target_) Then
                    If value.Length > maxTargetLength Then
                        value = value.Substring(0, maxTargetLength)
                    End If

                    Me.target_ = value
                    Me.WriteLink()
                End If
            End Set
        End Property
        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' ( Backing Field of <see cref="ShortcutFileInfo.Target"/> property. )
        ''' <para></para>
        ''' The full path of the target file or directory.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private target_ As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the command-line arguments of the target.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The command-line arguments for the target. Max Length is 1023 characters.")>
        <DisplayName("Target Arguments")>
        <Category("Shortcut")>
        <DefaultValue("")>
        Public Property TargetArguments As String
            Get
                Return Me.targetArguments_
            End Get
            Set(value As String)
                If (value <> Me.targetArguments_) Then
                    If value.Length > maxArgumentsLength Then
                        value = value.Substring(0, maxArgumentsLength)
                    End If
                    Me.targetArguments_ = value
                    Me.WriteLink()
                End If
            End Set
        End Property
        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' ( Backing Field of <see cref="ShortcutFileInfo.TargetArguments"/> property. )
        ''' <para></para>
        ''' The command-line arguments of the target.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private targetArguments_ As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the display name of the target file or directory.
        ''' <para></para>
        ''' Returns a empty string if the target does not exist.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The display name of the target file or directory. Or a empty string if the target does not exist.")>
        <DisplayName("Target Display Name")>
        <Category("Shortcut")>
        <DefaultValue("")>
        Public ReadOnly Property TargetDisplayName As String
            Get
                Dim shellItem As IShellItem = Nothing
                If Not String.IsNullOrWhiteSpace(Me.target_) Then
                    NativeMethods.SHCreateItemFromParsingName(Me.target_, IntPtr.Zero, GetType(IShellItem).GUID, shellItem)
                    Return shellItem?.GetDisplayName(ShellItemGetDisplayName.NormalDisplay).ToString()
                End If
                Return String.Empty
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the working directory of the target file or directory.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The working directory of the target file or directory. Max Length is 259 characters.")>
        <DisplayName("Working Directory")>
        <Category("Shortcut")>
        <DefaultValue("")>
        <Editor(GetType(FolderNameEditor), GetType(UITypeEditor))>
        Public Property WorkingDirectory As String
            Get
                Return Me.workingDirectory_
            End Get
            Set(value As String)
                If (value <> Me.workingDirectory_) Then
                    If value.Length > maxWorkingDirLength Then
                        value = value.Substring(0, maxWorkingDirLength)
                    End If
                    Me.workingDirectory_ = value
                    Me.WriteLink()
                End If
            End Set
        End Property
        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' ( Backing Field of <see cref="ShortcutFileInfo.WorkingDirectory"/> property. )
        ''' <para></para>
        ''' The working directory of the target file or directory.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private workingDirectory_ As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a value that determine whether the shortcut has an ITEM ID LIST pointing to a Windows Installer product.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("Determine whether the shortcut has an ITEM ID LIST pointing to a Windows Installer product.")>
        <DisplayName("Is Windows Installer Shortcut")>
        <Category("Shortcut")>
        <Browsable(False)>
        Public ReadOnly Property IsWindowsInstallerShortcut As Boolean
            Get
                Return Me.isWindowsInstallerShortcut_
            End Get
        End Property
        Private isWindowsInstallerShortcut_ As Boolean

#End Region

#Region " Other "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value indicating whether this <see cref="ShortcutFileInfo"/> is in 'view mode',
        ''' for example, being displayed in a <see cref="PropertyGrid"/> control.
        ''' <para></para>
        ''' The link file will not be modified while <see cref="ShortcutFileInfo.ViewMode"/> is <see langword="True"/>. 
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if this <see cref="ShortcutFileInfo"/> is displayed in a <see cref="PropertyGrid"/> control; 
        ''' otherwise, <see langword="False"/>.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Browsable(False)>
        <EditorBrowsable(EditorBrowsableState.Advanced)>
        Public Property ViewMode As Boolean
            Get
                Return Me.viewMode_
            End Get
            Set(value As Boolean)
                If (value <> Me.viewMode_) Then
                    Me.viewMode_ = value
                    Me.WriteLink()
                End If
            End Set
        End Property
        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' ( Backing Field of <see cref="ShortcutFileInfo.ViewMode"/> property. )
        ''' <para></para>
        ''' A value indicating whether this <see cref="ShortcutFileInfo"/> is in 'view mode',
        ''' for example, being displayed in a <see cref="PropertyGrid"/> control.
        ''' <para></para>
        ''' The link file will not be modified while <see cref="ShortcutFileInfo.ViewMode"/> is <see langword="True"/>. 
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private viewMode_ As Boolean

#End Region

#End Region

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Initializes a new instance of the <see cref="ShortcutFileInfo"/> class.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="filepath">
        ''' The fully qualified path of the new link file, or the relative file name. 
        ''' <para></para>
        ''' Do not end the path with the directory separator character.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <SecuritySafeCritical>
        Public Sub New(filePath As String)
            Me.OriginalPath = filePath
            Me.FullPath = Path.GetFullPath(filePath)

            If (Me.Exists) Then
                Me.ReadLink()
            End If
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Initializes a new instance of the <see cref="ShortcutFileInfo"/> class.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="file">
        ''' The link file. 
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <SecuritySafeCritical>
        Public Sub New(file As FileInfo)
            Me.New(file.FullName)
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Prevents a default instance of the <see cref="ShortcutFileInfo"/> class from being created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub New()
        End Sub

#End Region

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Refreshes the state of the object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shadows Sub Refresh()
            MyBase.Refresh()

            If (Me.Exists) Then
                Me.ReadLink()
            End If
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Creates the link file. It overwrites any existing file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Sub Create()
            Dim cShellLink As New CShellLink()
            Dim shellLink As IShellLinkW = DirectCast(cShellLink, IShellLinkW)
            Dim persistFile As IPersistFile = DirectCast(shellLink, IPersistFile)
            persistFile.Save(Me.FullName, True)

            Marshal.FinalReleaseComObject(persistFile)
            Marshal.FinalReleaseComObject(shellLink)
            Marshal.FinalReleaseComObject(cShellLink)

            Dim oldViewMode As Boolean = Me.viewMode_
            Me.viewMode_ = False
            Me.WriteLink()
            Me.viewMode_ = oldViewMode
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Deletes the link file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        <SecuritySafeCritical>
        Public Overrides Sub Delete()
            Dim file As New FileInfo(Me.FullName)
            file.Delete()
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Encrypts a link file so that only the account used to encrypt the file can decrypt it.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <ComVisible(False)>
        <DebuggerStepThrough>
        Public Sub Encrypt()
            Dim file As New FileInfo(Me.FullName)
            file.Encrypt()
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Decrypts a link file that was encrypted by the current account using the <see cref="ShortcutFileInfo.Encrypt"/> method.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <ComVisible(False)>
        <DebuggerStepThrough>
        Public Sub Decrypt()
            Dim file As New FileInfo(Me.FullName)
            file.Decrypt()
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Copies an existing link file to a new file, allowing the overwriting of an existing file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="destFileName">
        ''' The name of the new file to copy to.
        ''' </param>
        ''' 
        ''' <param name="overwrite">
        ''' <see langword="True"/> to allow an existing file to be overwritten; otherwise, <see langword="False"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A new link file, or an overwrite of an existing file if overwrite is true. 
        ''' <para></para>
        ''' If the file exists and overwrite is false, an <see cref="IOException"/> is thrown.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Function CopyTo(destFileName As String, overwrite As Boolean) As ShortcutFileInfo
            Return New ShortcutFileInfo(New FileInfo(Me.FullName).CopyTo(destFileName, overwrite))
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Moves the link file to a new location, providing the option to specify a new file name.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="destFileName">
        ''' The path to move the link file to, which can specify a different file name.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        <SecuritySafeCritical>
        Public Sub MoveTo(destFileName As String)
            Dim file As New FileInfo(Me.FullName)
            file.MoveTo(destFileName)
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Opens the link file in the specified mode.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="mode">
        ''' A <see cref="FileMode"/> constant specifying the mode (for example, Open or Append) in which to open the link file.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The link file opened in the specified mode, with read/write access and unshared.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Function Open(mode As FileMode) As FileStream
            Return New FileInfo(Me.FullName).Open(mode)
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Opens the link file in the specified mode with read, write, or read/write access.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="mode">
        ''' A <see cref="FileMode"/> constant specifying the mode (for example, Open or Append) in which to open the link file.
        ''' </param>
        ''' 
        ''' <param name="access">
        ''' A <see cref="FileAccess"/> constant specifying whether to open the link file with Read, Write, or ReadWrite file access.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The link file opened in the specified mode and access, and unshared.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Function Open(mode As FileMode, access As FileAccess) As FileStream
            Return New FileInfo(Me.FullName).Open(mode, access)
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Opens the link file in the specified mode with read, write, or read/write access and the specified sharing option.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="mode">
        ''' A <see cref="FileMode"/> constant specifying the mode (for example, Open or Append) in which to open the link file.
        ''' </param>
        ''' 
        ''' <param name="access">
        ''' A <see cref="FileAccess"/> constant specifying whether to open the link file with Read, Write, or ReadWrite file access.
        ''' </param>
        ''' 
        ''' <param name="share">
        ''' A <see cref="FileShare"/> constant specifying the type of access other <see cref="FileStream"/> objects have to this file.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The link file opened in the specified mode, access, and sharing options.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Function Open(mode As FileMode, access As FileAccess, share As FileShare) As FileStream
            Return New FileInfo(Me.FullName).Open(mode, access, share)
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Creates a read-only <see cref="FileStream"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A read-only unshared <see cref="FileStream"/> object for the existing link file.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Function OpenRead() As FileStream
            Return New FileInfo(Me.FullName).OpenRead()
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Creates a write-only <see cref="FileStream"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A write-only unshared <see cref="FileStream"/> object for the existing link file.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Function OpenWrite() As FileStream
            Return New FileInfo(Me.FullName).OpenWrite()
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a <see cref="FileSecurity"/> object that encapsulates the access control list (ACL) entries 
        ''' for the file described by the current <see cref="ShortCutFileInfo"/> object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A <see cref="FileSecurity"/> object that encapsulates the access control rules for the current link file.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Function GetAccessControl() As FileSecurity
            Return New FileInfo(Me.FullName).GetAccessControl()
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a <see cref="FileSecurity"/> object that encapsulates the specified type of access control list (ACL) entries 
        ''' for the file described by the current <see cref="ShortCutFileInfo"/> object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="includeSections">
        ''' One of the <see cref="AccessControlSections"/> values that specifies which group of access control entries to retrieve.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A <see cref="FileSecurity"/> object that encapsulates the access control rules for the current link file.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Function GetAccessControl(includeSections As AccessControlSections) As FileSecurity
            Return New FileInfo(Me.FullName).GetAccessControl(includeSections)
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Returns the shortcut' path as a string.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A string representing the shortcut's path.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Overrides Function ToString() As String
            Return Me.OriginalPath
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Determines whether the specified <see cref="ShortcutFileInfo"/> is equal to this <see cref="ShortcutFileInfo"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="file">
        ''' The <see cref="ShortcutFileInfo"/> to compare.
        ''' </param>        
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see langword="True"/> if the specified <see cref="ShortcutFileInfo"/> is equal to this <see cref="ShortcutFileInfo"/>; 
        ''' otherwise, <see langword="False"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shadows Function Equals(file As ShortcutFileInfo) As Boolean
            Return (Me = file)
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Resolves the target of a shortcut.
        ''' <para></para>
        ''' This is useful when the target of a link file is changed from a drive letter to another, for example.
        ''' <para></para>
        ''' If the target can't be resolved, an error message would be displayed.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="owner">
        ''' A handle used to show an error message If the target can't be resolved.
        ''' <para></para>
        ''' This value can be <see cref="IntPtr.Zero"/>
        ''' <para></para>
        ''' Add <see cref="IShellLinkResolveFlags.NoUI"/> flag to <see cref="flags"/> parameter 
        ''' to don't show any error message.
        ''' </param>
        ''' 
        ''' <param name="flags">
        ''' Flags that determine the resolve behavior.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Sub Resolve(owner As IntPtr, flags As IShellLinkResolveFlags)

            Dim lnk As New CShellLink()
            Dim lnkW As IShellLinkW
            Dim persistFile As IPersistFile = DirectCast(lnk, IPersistFile)

            persistFile.Load(Me.FullPath, 0)
            lnkW = DirectCast(lnk, IShellLinkW)
            lnkW.Resolve(owner, flags)

            Marshal.FinalReleaseComObject(persistFile)
            Marshal.FinalReleaseComObject(lnkW)
            Marshal.FinalReleaseComObject(lnk)

        End Sub

        Public Overrides Function GetHashCode() As Integer

            ' https://stackoverflow.com/a/4656890
            Return CInt(New With {
                Key .A = MyBase.FullName,
                Key .B = Me.description_,
                Key .C = Me.hotkey_,
                Key .D = Me.icon_,
                Key .E = Me.iconIndex_,
                Key .F = Me.target_,
                Key .G = Me.targetArguments_,
                Key .H = Me.windowState_,
                Key .I = Me.workingDirectory_
            }.GetHashCode() And &H7FFFFFFFL)

        End Function

#End Region

#Region " Private Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Reads the information from the link file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepperBoundary>
        Private Sub ReadLink()
            Dim arguments As New StringBuilder(capacity:=maxArgumentsLength, maxCapacity:=maxArgumentsLength)
            Dim description As New StringBuilder(capacity:=maxDescriptionLength, maxCapacity:=maxDescriptionLength)
            Dim hotkey As UShort
            Dim icon As New StringBuilder(capacity:=maxIconLength, maxCapacity:=maxIconLength)
            Dim iconIndex As Integer
            Dim idlist As IntPtr
            Dim target As New StringBuilder(capacity:=maxTargetLength, maxCapacity:=maxTargetLength)
            Dim windowstateNative As NativeWindowState = NativeWindowState.Normal
            Dim windowstate As ShortcutWindowState = ShortcutWindowState.Normal
            Dim workingDir As New StringBuilder(capacity:=maxWorkingDirLength, maxCapacity:=maxWorkingDirLength)

            Dim cShellLink As New CShellLink()
            Dim persistFile As IPersistFile = DirectCast(cShellLink, IPersistFile)
            persistFile.Load(Me.FullPath, 0)

            Dim shellLink As IShellLinkW = DirectCast(cShellLink, IShellLinkW)
            With shellLink
                .GetArguments(arguments, arguments.MaxCapacity)
                .GetHotkey(hotkey)
                .GetIDList(idlist)
                .GetIconLocation(icon, icon.MaxCapacity, iconIndex)
                .GetShowCmd(windowstateNative)
                windowstate = CType([Enum].ToObject(GetType(ShortcutWindowState), windowstateNative), ShortcutWindowState)
                .GetWorkingDirectory(workingDir, workingDir.MaxCapacity)
                Try
                    .GetDescription(description, description.MaxCapacity)
                Catch ex As Exception
                    ' Ignore.
                End Try
            End With

            Dim msiProductCode As New StringBuilder(39)
            Dim msiFeatureId As New StringBuilder(39)
            Dim msiComponentCode As New StringBuilder(39)
            Dim isMsiShortcut As Boolean = (NativeMethods.MsiGetShortcutTargetW(Me.FullPath, msiProductCode, msiFeatureId, msiComponentCode) = 0)
            Me.isWindowsInstallerShortcut_ = isMsiShortcut

            If isMsiShortcut Then
                Dim path As New StringBuilder(capacity:=1024)
                Dim size As Integer = path.Capacity
                Dim result As Integer = NativeMethods.MsiGetComponentPathW(msiProductCode.ToString(), msiComponentCode.ToString(), path, size)

                If result <= 0 Then
                    ' Values in "msi.h" header file:
                    '
                    ' INSTALLSTATE_NOTUSED      = -7,  // component disabled
                    ' INSTALLSTATE_BADCONFIG    = -6,  // configuration data corrupt
                    ' INSTALLSTATE_INCOMPLETE   = -5,  // installation suspended or in progress
                    ' INSTALLSTATE_SOURCEABSENT = -4,  // run from source, source is unavailable
                    ' INSTALLSTATE_MOREDATA     = -3,  // return buffer overflow
                    ' INSTALLSTATE_INVALIDARG   = -2,  // invalid function argument
                    ' INSTALLSTATE_UNKNOWN      = -1,  // unrecognized product or feature
                    ' INSTALLSTATE_BROKEN       =  0,  // broken
                    ' INSTALLSTATE_ADVERTISED   =  1,  // advertised feature
                    ' INSTALLSTATE_REMOVED      =  1,  // component being removed (action state, not settable)
                    ' INSTALLSTATE_ABSENT       =  2,  // uninstalled (or action state absent but clients remain)
                    ' INSTALLSTATE_LOCAL        =  3,  // installed on local drive
                    ' INSTALLSTATE_SOURCE       =  4,  // run from source, CD or net
                    ' INSTALLSTATE_DEFAULT      =  5,  // use default, local or source
                    Throw New Exception($"Cannot parse target of Windows Installer link. MSI Error Code: {result}")
                End If
                target = path

            Else
                ' SHGetNameFromIDList() can retrieve common file system paths, and CLSIDs/virtual folders.
                If (idlist = IntPtr.Zero) OrElse NativeMethods.SHGetNameFromIDList(idlist, ShellItemGetDisplayName.DesktopAbsoluteParsing, target) <> HResult.S_OK Then
                    target?.Clear()

                    ' IShellLinkW.GetPath() only can retrieve common file system paths.
                    shellLink.GetPath(target, target.Capacity, Nothing, IShellLinkGetPathFlags.RawPath)
                End If
            End If

            Marshal.FinalReleaseComObject(persistFile)
            Marshal.FinalReleaseComObject(shellLink)
            Marshal.FinalReleaseComObject(cShellLink)

            Me.description_ = description.ToString()
            Me.icon_ = icon.ToString()
            Me.iconIndex_ = iconIndex
            Me.target_ = target.ToString()
            Me.targetArguments_ = arguments.ToString()
            Me.windowState_ = windowstate
            Me.workingDirectory_ = workingDir.ToString()

            Dim keyModifier As ShortcutHotkeyModifier = CType(BitConverter.GetBytes(hotkey)(1), ShortcutHotkeyModifier)
            Dim keyAccesor As Keys = CType(BitConverter.GetBytes(hotkey)(0), Keys)
            Me.hotkey_ = Me.HotkeyToKeys(keyModifier, keyAccesor)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Writes the information to the link file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Private Sub WriteLink()
            If (Me.viewMode_) Then
                Exit Sub
            End If

            If Not Me.Exists Then
                Throw New FileNotFoundException($"The specified link file is not found: {Me.FullName}", Me.FullName)
            End If

            Dim isReadOnly As Boolean = Me.IsReadOnly
            If (isReadOnly) Then
                Me.Attributes = Me.Attributes And Not FileAttributes.ReadOnly
            End If

            Dim cShellLink As New CShellLink()
            Dim persistFile As IPersistFile = DirectCast(cShellLink, IPersistFile)
            persistFile.Load(Me.FullPath, 0)

            Dim shellLink As IShellLinkW = DirectCast(cShellLink, IShellLinkW)
            With shellLink
                If Not String.IsNullOrEmpty(Me.target_) Then
                    .SetPath(Me.target_) ' Will throw error if empty string.
                End If

                .SetArguments(Me.targetArguments_)
                .SetDescription(Me.description_)
                .SetHotkey(Me.KeysToHotkey(Me.hotkey_))
                .SetIconLocation(Me.icon_, Me.iconIndex_)
                .SetShowCmd(DirectCast(Me.windowState_, NativeWindowState))
                .SetWorkingDirectory(Me.workingDirectory_)
            End With

            Dim shellLinkPersistFile As IPersistFile = DirectCast(shellLink, IPersistFile)
            shellLinkPersistFile.Save(Me.FullPath, True)

            If Not String.IsNullOrEmpty(Me.target_) Then
                shellLinkPersistFile.SaveCompleted(Me.FullPath)
            End If

            If shellLinkPersistFile.IsDirty = 0 Then
                Throw New Exception(message:="Cannot save changes to link file, ensure that you are using an user account with granted permissions to write on this file.")
            End If

            Marshal.FinalReleaseComObject(shellLinkPersistFile)
            Marshal.FinalReleaseComObject(persistFile)
            Marshal.FinalReleaseComObject(shellLink)
            Marshal.FinalReleaseComObject(cShellLink)

            If isReadOnly Then
                Me.Attributes = Me.Attributes Or FileAttributes.ReadOnly
            End If
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Converts the shortcut hotkey to a <see cref="Keys"/> value.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="keyModifier">
        ''' The modifier key.
        ''' </param>
        ''' 
        ''' <param name="keyAccesor">
        ''' The accesor key.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The resulting <see cref="Keys"/> value.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Private Function HotkeyToKeys(keyModifier As ShortcutHotkeyModifier, keyAccesor As Keys) As Keys

            Dim key As Keys = keyAccesor

            If keyModifier.HasFlag(ShortcutHotkeyModifier.Alt) Then
                key = key Or Keys.Alt
            End If

            If keyModifier.HasFlag(ShortcutHotkeyModifier.Control) Then
                key = key Or Keys.Control
            End If

            If keyModifier.HasFlag(ShortcutHotkeyModifier.Shift) Then
                key = key Or Keys.Shift
            End If

            Return key

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Converts a <see cref="Keys"/> value to a valid shortcut hotkey.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="key">
        ''' The <see cref="Keys"/> value.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The resulting shortcut hotkey.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Private Function KeysToHotkey(key As Keys) As UShort

            Dim keyModifier As ShortcutHotkeyModifier = ShortcutHotkeyModifier.None
            Dim keyAccesor As Keys = CType(BitConverter.GetBytes(key)(0), Keys)

            If key.HasFlag(Keys.Alt) Then
                keyModifier = keyModifier Or ShortcutHotkeyModifier.Alt
            End If

            If key.HasFlag(Keys.Control) Then
                keyModifier = keyModifier Or ShortcutHotkeyModifier.Control
            End If

            If key.HasFlag(Keys.Shift) Then
                keyModifier = keyModifier Or ShortcutHotkeyModifier.Shift
            End If

            Return BitConverter.ToUInt16({CByte(keyAccesor), CByte(keyModifier)}, 0)

        End Function

#End Region

#Region " Operators "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Performs an implicit conversion from <see cref="ShortcutFileInfo"/> to <see cref="FileInfo"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="file">
        ''' The <see cref="ShortcutFileInfo"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The resulting <see cref="FileInfo"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Widening Operator CType(file As ShortcutFileInfo) As FileInfo
            Return New FileInfo(file.ToString())
        End Operator

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Performs an implicit conversion from <see cref="FileInfo"/> to <see cref="ShortcutFileInfo"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="file">
        ''' The <see cref="FileInfo"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The resulting <see cref="ShortcutFileInfo"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Widening Operator CType(file As FileInfo) As ShortcutFileInfo
            Return New ShortcutFileInfo(file)
        End Operator

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Determines whether the specified <see cref="ShortcutFileInfo"/> instances are equal.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="first">
        ''' The first <see cref="ShortcutFileInfo"/> to compare.
        ''' </param>
        ''' 
        ''' <param name="second">
        ''' The second <see cref="ShortcutFileInfo"/> to compare.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see langword="True"/> if the <see cref="ShortcutFileInfo"/> instances are considered equal; 
        ''' otherwise, <see langword="False"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Operator =(first As ShortcutFileInfo, second As ShortcutFileInfo) As Boolean

            Return (first.Attributes = second.Attributes) AndAlso
                   (first.CreationTime = second.CreationTime) AndAlso
                   (first.description_ = second.description_) AndAlso
                   (first.FullPath = second.FullPath) AndAlso
                   (first.hotkey_ = second.hotkey_) AndAlso
                   (first.icon_ = second.icon_) AndAlso
                   (first.iconIndex_ = second.iconIndex_) AndAlso
                   (first.LastAccessTime = second.LastAccessTime) AndAlso
                   (first.LastWriteTime = second.LastWriteTime) AndAlso
                   (first.target_ = second.target_) AndAlso
                   (first.targetArguments_ = second.targetArguments_) AndAlso
                   (first.windowState_ = second.windowState_) AndAlso
                   (first.workingDirectory_ = second.workingDirectory_)

        End Operator

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Determines whether the specified <see cref="ShortcutFileInfo"/> instances are not equal.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="first">
        ''' The first <see cref="ShortcutFileInfo"/> to compare.
        ''' </param>
        ''' 
        ''' <param name="second">
        ''' The second <see cref="ShortcutFileInfo"/> to compare.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see langword="True"/> if the <see cref="ShortcutFileInfo"/> instances are not equal; 
        ''' otherwise, <see langword="False"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Operator <>(first As ShortcutFileInfo, second As ShortcutFileInfo) As Boolean
            Return Not (first = second)
        End Operator

#End Region

    End Class

End Namespace

#End Region
