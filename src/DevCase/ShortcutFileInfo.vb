' ***********************************************************************
' Author   : ElektroStudios
' Modified : 14-November-2015
' ***********************************************************************

#Region " Public Members Summary "

#Region " Constructors "

' ShortcutInfo.New()

#End Region

#Region " Properties "


#End Region

#End Region

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security
Imports System.Security.AccessControl
Imports System.Text
Imports System.Xml.Serialization

Imports DevCase.Interop.Unmanaged.Win32.Enums
Imports DevCase.Interop.Unmanaged.Win32.Interfaces

#End Region

#Region " ShortcutFileInfo "

Namespace DevCase.Core.IO

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Provides information about a shortcut (.lnk) file.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    <Serializable>
    <XmlRoot(NameOf(ShortcutFileInfo))>
    <TypeConverter(GetType(ExpandableObjectConverter))>
    <Category(NameOf(ShortcutFileInfo))>
    <Description("Provides information about a shortcut (.lnk) file.")>
    <DefaultProperty(NameOf(ShortcutFileInfo.FullName))>
    Public NotInheritable Class ShortcutFileInfo : Inherits FileSystemInfo

#Region " Properties "

#Region " File Info "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the file name of the shortcut file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The name of the shortcut file.")>
        <DisplayName("File Name")>
        <Category("File Info")>
        Public Overrides ReadOnly Property Name As String
            Get
                Return Path.GetFileName(MyBase.FullName)
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the full path of the shortcut file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The full path of the shortcut file.")>
        <DisplayName("Full Path")>
        <Category("File Info")>
        Public Overrides ReadOnly Property FullName As String
            Get
                Return MyBase.FullName
            End Get
        End Property

#End Region

#Region " File Info Extended "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the file attributes of the current shortcut file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The file attributes.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The file attributes of the current shortcut file.")>
        <DisplayName("Attributes")>
        <Category("File Info Extended")>
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
        ''' Gets or sets the creation time of the current shortcut file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The creation time.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The creation time of the current shortcut file.")>
        <DisplayName("Creation Time")>
        <Category("File Info Extended")>
        <Browsable(True)>
        Public Shadows Property CreationTime As Date
            Get
                Return MyBase.CreationTime
            End Get
            Set(value As Date)
                MyBase.CreationTime = value
            End Set
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the time the current shortcut file was last accessed.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The last access time.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The time the current shortcut file was last accessed.")>
        <DisplayName("Last Access Time")>
        <Category("File Info Extended")>
        <Browsable(True)>
        Public Shadows Property LastAccessTime As Date
            Get
                Return MyBase.LastAccessTime
            End Get
            Set(value As Date)
                MyBase.LastAccessTime = value
            End Set
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the time the current shortcut file was last written to.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The last write time.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The time the current shortcut file was last written to.")>
        <DisplayName("Last Write Time")>
        <Category("File Info Extended")>
        <Browsable(True)>
        Public Shadows Property LastWriteTime As Date
            Get
                Return MyBase.LastWriteTime
            End Get
            Set(value As Date)
                MyBase.LastWriteTime = value
            End Set
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the file size, in bytes, of the current shortcut file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The file attributes.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The file size, in bytes, of the current shortcut file.")>
        <DisplayName("Length")>
        <Category("File Info Extended")>
        <Browsable(True)>
        Public ReadOnly Property Length As Long
            Get
                Return New FileInfo(Me.FullName).Length
            End Get
        End Property

#End Region

#Region " File Info Extended (not browsable)"

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines if the current shortcut file is read-only.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if the current shortcut file is read-only; otherwise, <see langword="False"/>.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
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
        ''' Gets a value indicating whether the shortcut file exists.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if the shortcut file exists; otherwise, <see langword="False"/>.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Browsable(False)>
        Public Overrides ReadOnly Property Exists As Boolean
            Get
                Return File.Exists(Me.FullName)
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets an instance of the parent directory of the shortcut file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
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
        <Browsable(False)>
        Public ReadOnly Property DirectoryName As String
            Get
                Return New FileInfo(Me.FullName).DirectoryName
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the string representing the extension part of the shortcut file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The extension.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Browsable(False)>
        Public Shadows ReadOnly Property Extension As String
            Get
                Return MyBase.Extension
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the creation time, in coordinated universal time (UTC) of the current shortcut file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The creation time UTC.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Browsable(False)>
        Public Shadows Property CreationTimeUtc As Date
            Get
                Return MyBase.CreationTimeUtc
            End Get
            Set(value As Date)
                MyBase.CreationTimeUtc = value
            End Set
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the time, in coordinated universal time (UTC), the current shortcut file was last accessed.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The last access time UTC.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Browsable(False)>
        Public Shadows Property LastAccessTimeUtc As Date
            Get
                Return MyBase.LastAccessTimeUtc
            End Get
            Set(value As Date)
                MyBase.LastAccessTimeUtc = value
            End Set
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the time, in coordinated universal time (UTC), the current shortcut file was last written to.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The last write time UTC.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Browsable(False)>
        Public Shadows Property LastWriteTimeUtc As Date
            Get
                Return MyBase.LastWriteTimeUtc
            End Get
            Set(value As Date)
                MyBase.LastWriteTimeUtc = value
            End Set
        End Property

#End Region

#Region " Shortcut "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the shortcut description.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The shortcut description.")>
        <DisplayName("Description")>
        <Category("Shortcut")>
        Public Property Description As String
            Get
                Return Me.description_
            End Get
            Set(value As String)
                If (value <> Me.description_) Then
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
        ''' Gets or sets the full path to the icon file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The full path to the icon file.")>
        <DisplayName("Icon")>
        <Category("Shortcut")>
        Public Property Icon As String
            Get
                Return Me.icon_
            End Get
            Set(value As String)
                If (value <> Me.icon_) Then
                    Me.icon_ = value
                    Me.WriteLink()
                End If
            End Set
        End Property
        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' ( Backing Field of <see cref="ShortcutFileInfo.Icon"/> property. )
        ''' <para></para>
        ''' The full path to the icon file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private icon_ As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the index of the image to use for the icon file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The index of the image to use for the icon file.")>
        <DisplayName("Icon Index")>
        <Category("Shortcut")>
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
        ''' The index of the image to use for the icon file.
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
        ''' The window state for the target file or directory.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private windowState_ As ShortcutWindowState

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the full path to the target file or directory.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The full path to the target file or directory.")>
        <DisplayName("Target")>
        <Category("Shortcut")>
        Public Property Target As String
            Get
                Return Me.target_
            End Get
            Set(value As String)
                If (value <> Me.target_) Then
                    Me.target_ = value
                    Me.WriteLink()
                End If
            End Set
        End Property
        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' ( Backing Field of <see cref="ShortcutFileInfo.Target"/> property. )
        ''' <para></para>
        ''' The full path to the target file or directory.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private target_ As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the command-line arguments for a target executable file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The command-line arguments for a target executable file.")>
        <DisplayName("Target Arguments")>
        <Category("Shortcut")>
        Public Property TargetArguments As String
            Get
                Return Me.targetArguments_
            End Get
            Set(value As String)
                If (value <> Me.targetArguments_) Then
                    Me.targetArguments_ = value
                    Me.WriteLink()
                End If
            End Set
        End Property
        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' ( Backing Field of <see cref="ShortcutFileInfo.TargetArguments"/> property. )
        ''' <para></para>
        ''' The command-line arguments for a target executable file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private targetArguments_ As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the working directory of the target file or directory.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("The working directory of the target file or directory.")>
        <DisplayName("Working Directory")>
        <Category("Shortcut")>
        Public Property WorkingDirectory As String
            Get
                Return Me.workingDirectory_
            End Get
            Set(value As String)
                If (value <> Me.workingDirectory_) Then
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

#End Region

#End Region

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Initializes a new instance of the <see cref="ShortcutFileInfo"/> class.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="filepath">
        ''' The fully qualified path of the new shortcut file, or the relative file name. 
        ''' <para></para>
        ''' Do not end the path with the directory separator character.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <SecuritySafeCritical>
        Public Sub New(ByVal filePath As String)
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
        ''' The shortcut file. 
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <SecuritySafeCritical>
        Public Sub New(ByVal file As FileInfo)
            Me.New(file.FullName)
        End Sub

#End Region

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Creates the shortcut file.
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
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Deletes the shortcut file.
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
        ''' Encrypts a shortcut file so that only the account used to encrypt the file can decrypt it.
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
        ''' Decrypts a shortcut file that was encrypted by the current account using the <see cref="ShortcutFileInfo.Encrypt"/> method.
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
        ''' Copies an existing shortcut file to a new file, allowing the overwriting of an existing file.
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
        ''' A new shortcut file, or an overwrite of an existing file if overwrite is true. 
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
        ''' Moves the shortcut file to a new location, providing the option to specify a new file name.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="destFileName">
        ''' The path to move the shortcut file to, which can specify a different file name.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        <SecuritySafeCritical>
        Public Sub MoveTo(ByVal destFileName As String)
            Dim file As New FileInfo(Me.FullName)
            file.MoveTo(destFileName)
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Opens the shortcut file in the specified mode.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="mode">
        ''' A <see cref="FileMode"/> constant specifying the mode (for example, Open or Append) in which to open the shortcut file.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The shortcut file opened in the specified mode, with read/write access and unshared.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Function Open(ByVal mode As FileMode) As FileStream
            Return New FileInfo(Me.FullName).Open(mode)
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Opens the shortcut file in the specified mode with read, write, or read/write access.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="mode">
        ''' A <see cref="FileMode"/> constant specifying the mode (for example, Open or Append) in which to open the shortcut file.
        ''' </param>
        ''' 
        ''' <param name="access">
        ''' A <see cref="FileAccess"/> constant specifying whether to open the shortcut file with Read, Write, or ReadWrite file access.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The shortcut file opened in the specified mode and access, and unshared.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Function Open(mode As FileMode, access As FileAccess) As FileStream
            Return New FileInfo(Me.FullName).Open(mode, access)
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Opens the shortcut file in the specified mode with read, write, or read/write access and the specified sharing option.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="mode">
        ''' A <see cref="FileMode"/> constant specifying the mode (for example, Open or Append) in which to open the shortcut file.
        ''' </param>
        ''' 
        ''' <param name="access">
        ''' A <see cref="FileAccess"/> constant specifying whether to open the shortcut file with Read, Write, or ReadWrite file access.
        ''' </param>
        ''' 
        ''' <param name="share">
        ''' A <see cref="FileShare"/> constant specifying the type of access other <see cref="FileStream"/> objects have to this file.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The shortcut file opened in the specified mode, access, and sharing options.
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
        ''' A read-only unshared <see cref="FileStream"/> object for the existing shortcut file.
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
        ''' A write-only unshared <see cref="FileStream"/> object for the existing shortcut file.
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
        ''' A <see cref="FileSecurity"/> object that encapsulates the access control rules for the current shortcut file.
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
        ''' A <see cref="FileSecurity"/> object that encapsulates the access control rules for the current shortcut file.
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

#End Region

#Region " Private Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Reads the information from the shortcut file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Private Sub ReadLink()

            Dim arguments As New StringBuilder(260)
            Dim description As New StringBuilder(260)
            Dim hotkey As UShort
            Dim icon As New StringBuilder(260)
            Dim iconIndex As Integer
            Dim target As New StringBuilder(260)
            Dim windowstate As ShortcutWindowState = ShortcutWindowState.Normal
            Dim workingDir As New StringBuilder(260)

            Dim cShellLink As New CShellLink()
            Dim persistFile As IPersistFile = DirectCast(cShellLink, IPersistFile)
            persistFile.Load(Me.FullPath, 0)

            Dim shellLink As IShellLinkW = DirectCast(cShellLink, IShellLinkW)
            With shellLink
                .GetArguments(arguments, arguments.Capacity)
                .GetDescription(description, description.Capacity)
                .GetHotkey(hotkey)
                .GetIconLocation(icon, icon.Capacity, iconIndex)
                .GetPath(target, target.Capacity, Nothing, IShellLinkGetPathFlags.UncPriority)
                .GetShowCmd(DirectCast(windowstate, NativeWindowState))
                .GetWorkingDirectory(workingDir, workingDir.Capacity)
            End With

            Marshal.FinalReleaseComObject(persistFile)
            Marshal.FinalReleaseComObject(shellLink)
            Marshal.FinalReleaseComObject(cShellLink)

            Me.targetArguments_ = arguments.ToString()
            Me.description_ = description.ToString()
            Me.icon_ = icon.ToString()
            Me.iconIndex_ = iconIndex
            Me.target_ = target.ToString()
            Me.windowState_ = windowstate
            Me.workingDirectory_ = workingDir.ToString()

            Dim keyModifier As ShortcutHotkeyModifier = CType(BitConverter.GetBytes(hotkey)(1), ShortcutHotkeyModifier)
            Dim keyAccesor As Keys = CType(BitConverter.GetBytes(hotkey)(0), Keys)
            Me.hotkey_ = Me.HotkeyToKeys(keyModifier, keyAccesor)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Writes the information to the shortcut file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Private Sub WriteLink()
            If Not Me.Exists Then
                Throw New FileNotFoundException($"The specified shortcut file is not found: {Me.FullName}", Me.FullName)
            End If

            If (Me.IsReadOnly) Then
                Throw New Exception($"The specified shortcut file is read-only: {Me.FullName}")
            End If

            Dim cShellLink As New CShellLink()
            Dim persistFile As IPersistFile = DirectCast(cShellLink, IPersistFile)
            persistFile.Load(Me.FullPath, 0)

            Dim shellLink As IShellLinkW = DirectCast(cShellLink, IShellLinkW)
            If Not String.IsNullOrEmpty(Me.target_) Then
                shellLink.SetPath(Me.target_)
            End If
            With shellLink
                .SetArguments(Me.targetArguments_)
                .SetDescription(Me.description_)
                .SetWorkingDirectory(Me.workingDirectory_)
                .SetIconLocation(Me.icon_, Me.iconIndex_)
                .SetHotkey(Me.KeysToHotkey(Me.hotkey_))
                .SetShowCmd(DirectCast(Me.windowState_, NativeWindowState))
            End With

            Dim shellLinkPersistFile As IPersistFile = DirectCast(shellLink, IPersistFile)
            shellLinkPersistFile.Save(Me.FullPath, True)

            If Not String.IsNullOrEmpty(Me.target_) Then
                shellLinkPersistFile.SaveCompleted(Me.FullPath)
            End If

            Marshal.FinalReleaseComObject(shellLinkPersistFile)
            Marshal.FinalReleaseComObject(persistFile)
            Marshal.FinalReleaseComObject(shellLink)
            Marshal.FinalReleaseComObject(cShellLink)
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
        Private Function HotkeyToKeys(ByVal keyModifier As ShortcutHotkeyModifier, ByVal keyAccesor As Keys) As Keys

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
        Private Function KeysToHotkey(ByVal key As Keys) As UShort

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

    End Class

End Namespace

#End Region
