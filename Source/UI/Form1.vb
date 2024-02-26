' ***********************************************************************
' Author   : ElektroStudios
' Modified : 29-June-2019
' ***********************************************************************

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.ComponentModel
Imports System.Globalization
Imports System.IO
Imports System.Threading

Imports DevCase.Core.Application.Tools
Imports DevCase.Core.Application.UserInterface
Imports DevCase.Core.Extensions
Imports DevCase.Core.IO
Imports DevCase.Core.IO.Tools
Imports DevCase.Core.Imaging.Tools

#End Region

#Region " Form1 "

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' The main application Form.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
Friend NotInheritable Class Form1 : Inherits Form

#Region " Private Fields "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' The current <see cref="ShortcutFileInfo"/> instance that is loaded in the <see cref="Form1.PropertyGrid1"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Private currentShortcut As ShortcutFileInfo
    
    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' The current <see cref="Be.Windows.Forms.DynamicFileByteProvider"/> instance that is loaded in the <see cref="Form1.HexBox1"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Private currentFileByteProvider As Be.Windows.Forms.DynamicFileByteProvider
 
#End Region

#Region " Constructors "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Initializes a new instance of the <see cref="Form1"/> class.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Public Sub New()
        MyClass.InitializeComponent()

        Dim ci As CultureInfo = CultureInfo.CreateSpecificCulture("en-US")
        Thread.CurrentThread.CurrentCulture = ci
        Thread.CurrentThread.CurrentUICulture = ci
        ' Set file dialogs language too.
        CultureUtil.SetProcessPreferredUILanguages(ci.Name)
    End Sub

#End Region

#Region " Event Handlers "

#Region " Form "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="Form.Load"/> event of the <see cref="Form1"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.ToolStripStatusLabelIcon.Text = ""
        Me.ToolStripStatusLabelFileName.Text = ""
        Me.ToolStripComboBoxFontSize.SelectedItem = CStr(My.Settings.FontSize)

        Me.LoadVisualTheme()

        ' Load file from command-line arguments.
        If My.Application.CommandLineArgs.Any()
            Dim data As New DataObject(DataFormats.FileDrop, My.Application.CommandLineArgs.ToArray())
            Dim args As new DragEventArgs(data, Nothing, 0, 0, Nothing, DragDropEffects.Copy)
            Me.PropertyGrid1_DragDrop(Me, args)
        End If

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="Form.Shown"/> event of the <see cref="Form1"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Dim minWidth As Integer = (From item As ToolStripMenuItem In Me.MenuStrip1.Items.Cast(Of ToolStripMenuItem)
                                   Select item.Width + item.Padding.Horizontal).Sum()

        Me.MinimumSize = New Size(minWidth, Me.Height)
    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="Form.Resize"/> event of the <see cref="Form1"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize

        Me.AdjustPropertyGridSplitter()

    End Sub

#End Region

#Region " Tab Control "
    
    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="Manina.Windows.Forms.PagedControl.PageChanged"/> event of the <see cref="Form1.TabControl1"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="Manina.Windows.Forms.PageChangedEventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub TabControl1_TabIndexChanged(sender As Object, e As Manina.Windows.Forms.PageChangedEventArgs) Handles TabControl1.PageChanged
        
        Me.AdjustPropertyGridSplitter()

    End Sub
    
#End Region

#Region " PropertyGrid "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="Control.DragDrop"/> event of the <see cref="Form1.PropertyGrid1"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="DragEventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub PropertyGrid1_DragDrop(sender As Object, e As DragEventArgs) Handles PropertyGrid1.DragDrop

        If (e.Effect = DragDropEffects.Copy) Then
            Dim filePath As String = DirectCast(e.Data.GetData(DataFormats.FileDrop), String()).Single()
            Dim fi As New FileInfo(filePath)
            If (fi.Exists) AndAlso (fi.Extension.Equals(".lnk", StringComparison.OrdinalIgnoreCase)) Then
                Me.LoadShortcutInPropertyGrid(filePath)
            End If
        End If

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="Control.DragEnter"/> event of the <see cref="Form1.PropertyGrid1"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="DragEventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub PropertyGrid1_DragEnter(sender As Object, e As DragEventArgs) Handles PropertyGrid1.DragEnter

        If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
            Dim filePaths() As String = DirectCast(e.Data.GetData(DataFormats.FileDrop), String())
            If (filePaths.Length = 1) AndAlso Path.GetExtension(filePaths(0)).Equals(".lnk", StringComparison.OrdinalIgnoreCase) Then
                e.Effect = DragDropEffects.Copy
            End If
        End If

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="PropertyGrid.PropertyValueChanged"/> event of the <see cref="Form1.PropertyGrid1"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="PropertyValueChangedEventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub PropertyGrid1_PropertyValueChanged(sender As Object, e As PropertyValueChangedEventArgs) Handles PropertyGrid1.PropertyValueChanged
        ' Force refresh of ShortcutFileInfo properties.
        Me.PropertyGrid1.Refresh()
    End Sub

#End Region
    
#Region " Hexadecimal Box (HexBox) "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="Control.DragDrop"/> event of the <see cref="Form1.HexBox1"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="DragEventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub HexBox1_DragDrop(sender As Object, e As DragEventArgs) Handles HexBox1.DragDrop
        
        If (e.Effect = DragDropEffects.Copy) Then
            Dim filePath As String = DirectCast(e.Data.GetData(DataFormats.FileDrop), String()).Single()
            Dim fi As New FileInfo(filePath)
            If (fi.Exists) AndAlso (fi.Extension.Equals(".lnk", StringComparison.OrdinalIgnoreCase)) Then
                Me.LoadShortcutInPropertyGrid(filePath)
            End If
        End If

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="Control.DragEnter"/> event of the <see cref="Form1.HexBox1"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="DragEventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub HexBox1_DragEnter(sender As Object, e As DragEventArgs) Handles HexBox1.DragEnter

        If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
            Dim filePaths() As String = DirectCast(e.Data.GetData(DataFormats.FileDrop), String())
            If (filePaths.Length = 1) AndAlso Path.GetExtension(filePaths(0)).Equals(".lnk", StringComparison.OrdinalIgnoreCase) Then
                e.Effect = DragDropEffects.Copy
            End If
        End If

    End Sub

#End Region

#Region " Menu Strip "

#Region " File Menu "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.Click"/> event of the <see cref="Form1.NewToolStripMenuItem"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click

        Using dlg As New SaveFileDialog()
            dlg.FileName = "New Shortcut.lnk"
            dlg.DefaultExt = "lnk"
            dlg.DereferenceLinks = False
            dlg.Filter = "Shortcut files (*.lnk)|*.lnk"
            dlg.FilterIndex = 1
            dlg.RestoreDirectory = True
            dlg.SupportMultiDottedExtensions = True
            dlg.Title = "Select a destination to save the shortcut file..."

            If dlg.ShowDialog() = DialogResult.OK Then
                Try
                    File.Delete(dlg.FileName)
                Catch ex As Exception
                    MessageBox.Show(Me, "Can't overwrite file.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                Dim newShortcut As New ShortcutFileInfo(dlg.FileName) With {.ViewMode = True}
                newShortcut.Create()

                Me.LoadShortcutInPropertyGrid(newShortcut.FullName)
            End If
        End Using

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.Click"/> event of the <see cref="Form1.OpenToolStripMenuItem"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click

        Using dlg As New OpenFileDialog()
            dlg.DefaultExt = "lnk"
            dlg.DereferenceLinks = False
            dlg.Filter = "Shortcut files (*.lnk)|*.lnk"
            dlg.FilterIndex = 1
            dlg.Multiselect = False
            dlg.RestoreDirectory = True
            dlg.SupportMultiDottedExtensions = True
            dlg.Title = "Select a shortcut file to open..."

            If dlg.ShowDialog = DialogResult.OK Then
                Me.LoadShortcutInPropertyGrid(dlg.FileName)
            End If

        End Using

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.Click"/> event of the <see cref="Form1.SaveToolStripMenuItem"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        Me.CurrentShortcut.Create()
        Me.PropertyGrid1.Refresh()
        Me.LoadShortcutInHexBox(Me.CurrentShortcut.FullName)
    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.Click"/> event of the <see cref="Form1.SaveAsToolStripMenuItem"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click

        Using dlg As New SaveFileDialog()
            dlg.FileName = Me.CurrentShortcut.FullName
            dlg.DefaultExt = "lnk"
            dlg.DereferenceLinks = False
            dlg.Filter = "Shortcut files (*.lnk)|*.lnk"
            dlg.FilterIndex = 1
            dlg.RestoreDirectory = True
            dlg.SupportMultiDottedExtensions = True
            dlg.Title = "Select a destination to save the shortcut file..."

            If dlg.ShowDialog() = DialogResult.OK Then

                Dim dstShortcut As New ShortcutFileInfo(dlg.FileName) With {.ViewMode = True}
                With dstShortcut
                    .Description = Me.CurrentShortcut.Description
                    .Hotkey = Me.CurrentShortcut.Hotkey
                    .Icon = Me.CurrentShortcut.Icon
                    .IconIndex = Me.CurrentShortcut.IconIndex
                    .Target = Me.CurrentShortcut.Target
                    .TargetArguments = Me.CurrentShortcut.TargetArguments
                    .WindowState = Me.CurrentShortcut.WindowState
                    .WorkingDirectory = Me.CurrentShortcut.WorkingDirectory

                    .Create()

                    .Attributes = Me.CurrentShortcut.Attributes
                    .CreationTime = Me.CurrentShortcut.CreationTime
                    .LastAccessTime = Me.CurrentShortcut.LastAccessTime
                    .LastWriteTime = Me.CurrentShortcut.LastWriteTime
                End With

            End If
        End Using

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.Click"/> event of the <see cref="Form1.CloseToolStripMenuItem"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        
        Me.HexBox1.ByteProvider = Nothing
        If Me.currentFileByteProvider IsNot Nothing
            Me.currentFileByteProvider.Dispose()
            Me.currentFileByteProvider = Nothing
        End If

        Me.PropertyGrid1.SelectedObject = Nothing
        Me.ToolStripStatusLabelFileName.Image = Nothing
        Me.ToolStripStatusLabelFileName.Text = ""

        Me.SaveToolStripMenuItem.Enabled = False
        Me.SaveAsToolStripMenuItem.Enabled = False
        Me.CloseToolStripMenuItem.Enabled = False

        Me.PropertyGrid1.ContextMenuStrip = Nothing
        Me.StatusStrip1.ContextMenuStrip = Nothing
    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.Click"/> event of the <see cref="Form1.ExitToolStripMenuItem"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

#End Region

#Region " Settings Menu "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripComboBox.SelectedIndexChanged"/> event of the <see cref="Form1.ToolStripComboBoxFontSize"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub ToolStripComboBoxFontSize_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ToolStripComboBoxFontSize.SelectedIndexChanged

        Dim sz As Single = CSng(DirectCast(sender, ToolStripComboBox).SelectedItem)
        My.Settings.FontSize = sz
        Me.LoadFontSize()

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.Click"/> event of the <see cref="Form1.DefaultToolStripMenuItem"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub DefaultToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DefaultToolStripMenuItem.Click
        My.Settings.VisualThemeIndex = 0
        Me.LoadVisualTheme()
    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.Click"/> event of the <see cref="Form1.DarkToolStripMenuItem"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub DarkToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DarkToolStripMenuItem.Click
        My.Settings.VisualThemeIndex = 1
        Me.LoadVisualTheme()
    End Sub

#End Region

#Region " About Menu "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.Click"/> event of the <see cref="Form1.AboutToolStripMenuItem"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        My.Forms.AboutBox1.ShowDialog()
    End Sub

#End Region

#End Region

#Region " PropertyGrid Context Menu "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripDropDown.Opening"/> event of the <see cref="Form1.ContextMenuStrip1"/> menu.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="CancelEventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub ContextMenuStrip1_Opening(sender As Object, e As CancelEventArgs) Handles ContextMenuStrip1.Opening

        Me.OpenShortcutMenuItem.Enabled = Me.CurrentShortcut.Exists
        Me.ViewShortcutMenuItem.Enabled = Me.OpenShortcutMenuItem.Enabled

        Me.OpenTargetMenuItem.Enabled = File.Exists(Me.CurrentShortcut.Target) OrElse Directory.Exists(Me.CurrentShortcut.Target)
        Me.OpenTargetWithArgsMenuItem.Enabled = File.Exists(Me.CurrentShortcut.Target) AndAlso Not String.IsNullOrWhiteSpace(Me.CurrentShortcut.TargetArguments)
        Me.ViewTargetMenuItem.Enabled = Me.OpenTargetMenuItem.Enabled

        Me.ViewWorkingDirectoryMenuItem.Enabled = Directory.Exists(Me.CurrentShortcut.WorkingDirectory)
        Me.ViewIconMenuItem.Enabled = File.Exists(Me.CurrentShortcut.Icon)

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.Click"/> event of the <see cref="Form1.OpenShortcutMenuItem"/> menu item.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub OpenShortcutMenuItem_Click(sender As Object, e As EventArgs) Handles OpenShortcutMenuItem.Click

        Try
            Process.Start(Me.CurrentShortcut.FullName)

        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.Click"/> event of the <see cref="Form1.OpenTargetMenuItem"/> menu item.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub OpenTargetMenuItem1_Click(sender As Object, e As EventArgs) Handles OpenTargetMenuItem.Click

        Dim target As String = Me.CurrentShortcut.Target

        If File.Exists(target) Then
            Try
                Process.Start(target)
                Exit Sub

            Catch ex As Exception
                MessageBox.Show(Me, ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)

            End Try
        End If

        If Directory.Exists(target) Then
            Try
                FileUtil.InternalOpenInExplorer(target)
                Exit Sub

            Catch ex As Exception
                MessageBox.Show(Me, ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)

            End Try
        End If

        MessageBox.Show(Me, "Can't find the shortcut target.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.Click"/> event of the <see cref="Form1.OpenTargetWithArgsMenuItem"/> menu item.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub OpenTargetWithArgsMenuItem_Click(sender As Object, e As EventArgs) Handles OpenTargetWithArgsMenuItem.Click

        Try
            Process.Start(Me.CurrentShortcut.Target, Me.CurrentShortcut.TargetArguments)
            Exit Sub

        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.Click"/> event of the <see cref="Form1.ViewShortcutMenuItem"/> menu item.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub ViewShortcutMenuItem_Click(sender As Object, e As EventArgs) Handles ViewShortcutMenuItem.Click

        Try
            FileUtil.InternalOpenInExplorer(Me.CurrentShortcut.FullName)

        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.Click"/> event of the <see cref="Form1.ViewTargetMenuItem"/> menu item.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub ViewTargetMenuItem_Click(sender As Object, e As EventArgs) Handles ViewTargetMenuItem.Click

        Try
            FileUtil.InternalOpenInExplorer(Me.CurrentShortcut.Target)

        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.Click"/> event of the <see cref="Form1.ViewWorkingDirectoryMenuItem"/> menu item.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub ViewWorkingDirectoryMenuItem_Click(sender As Object, e As EventArgs) Handles ViewWorkingDirectoryMenuItem.Click

        Try
            FileUtil.InternalOpenInExplorer(Me.CurrentShortcut.WorkingDirectory)

        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.Click"/> event of the <see cref="Form1.ViewIconMenuItem"/> menu item.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="EventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub ViewIconMenuItem_Click(sender As Object, e As EventArgs) Handles ViewIconMenuItem.Click

        Try
            FileUtil.InternalOpenInExplorer(Me.CurrentShortcut.Icon)

        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

#End Region

#End Region

#Region " Private Methods "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Loads the saved font size for the user-interface.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub LoadFontSize()
        Me.MenuStrip1.Font = New Font(Me.MenuStrip1.Font.FontFamily, My.Settings.FontSize, Me.MenuStrip1.Font.Style)
        Me.StatusStrip1.Font = New Font(Me.StatusStrip1.Font.FontFamily, My.Settings.FontSize, Me.StatusStrip1.Font.Style)
        Me.ToolStripStatusLabelIcon.Font = New Font(Me.ToolStripStatusLabelIcon.Font.FontFamily, My.Settings.FontSize, Me.ToolStripStatusLabelIcon.Font.Style)
        Me.ToolStripStatusLabelFileName.Font = New Font(Me.ToolStripStatusLabelFileName.Font.FontFamily, My.Settings.FontSize, Me.ToolStripStatusLabelFileName.Font.Style)
        Me.PropertyGrid1.Font = New Font(Me.PropertyGrid1.Font.FontFamily, My.Settings.FontSize, Me.PropertyGrid1.Font.Style)
    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Loads the saved visual theme for the user-interface.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub LoadVisualTheme()
        Select Case My.Settings.VisualThemeIndex

            Case 0 ' Default theme
                Me.SetVisualStyle(VisualStyle.Default, True)
                My.Forms.AboutBox1.SetVisualStyle(VisualStyle.Default, True)

                Me.DefaultToolStripMenuItem.Checked = True
                Me.DarkToolStripMenuItem.Checked = False

            Case 1 ' Dark theme
                Me.SetVisualStyle(VisualStyle.VisualStudioDark, True)
                My.Forms.AboutBox1.SetVisualStyle(VisualStyle.VisualStudioDark, True)

                Me.DefaultToolStripMenuItem.Checked = False
                Me.DarkToolStripMenuItem.Checked = True

            Case Else
                ' Do nothing.

        End Select
    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Updates the MRU menu items.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="shortcut">
    ''' The source <see cref="ShortcutFileInfo"/> to add at the top of the MRU items.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub UpdateMruItems(shortcut As ShortcutFileInfo)

        ' Search and remove existing MRU item.
        For Each item As ToolStripMenuItem In Me.RecentToolStripMenuItem.DropDown.Items
            If CStr(item.Tag).Equals(shortcut.FullName, StringComparison.OrdinalIgnoreCase) Then
                item.Image?.Dispose()
                item.Dispose()
                Exit For
            End If
        Next

        ' Keep maximum capacity.
        Const maxMruCapacity As Integer = 10
        If (Me.RecentToolStripMenuItem.DropDown.Items.Count = maxMruCapacity) Then
            Me.RecentToolStripMenuItem.DropDown.Items(maxMruCapacity - 1).Dispose()
        End If

        ' Retrieve shortcut's file icon.
        Dim ico As Icon = Nothing
        Try
            ico = ImageUtil.ExtractIconFromFile(shortcut.FullName)
        Catch
        End Try

        ' Add new MRU item.
        Dim newItem As New ToolStripMenuItem With {
            .Text = shortcut.Name,
            .Image = ico?.ToBitmap(),
            .Tag = shortcut.FullName
        }
        Me.RecentToolStripMenuItem.DropDown.Items.Insert(0, newItem)

        AddHandler newItem.Click, Sub(sender As Object, e As EventArgs)
                                      Dim fullPath As String = CStr(DirectCast(sender, ToolStripMenuItem).Tag)
                                      Me.LoadShortcutInPropertyGrid(fullPath)
                                  End Sub

        ico?.Dispose()

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Loads a shortcut file into the <see cref="Form1.PropertyGrid1"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="filePath">
    ''' The shortcut file (.lnk) path.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub LoadShortcutInPropertyGrid(filePath As String)

        Me.CurrentShortcut = New ShortcutFileInfo(filePath) With {.ViewMode = True}
        If Not Me.CurrentShortcut.Exists Then
            MessageBox.Show(Me, $"The lnk file does not exist: {Me.CurrentShortcut.FullName}", Me.Name, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Me.AdjustPropertyGridSplitter()
        Me.PropertyGrid1.SelectedObject = Me.CurrentShortcut
        Me.ToolStripStatusLabelFileName.Text = filePath

        Me.RecentToolStripMenuItem.Enabled = True
        Me.SaveToolStripMenuItem.Enabled = True
        Me.SaveAsToolStripMenuItem.Enabled = True
        Me.CloseToolStripMenuItem.Enabled = True

        Me.UpdateMruItems(Me.CurrentShortcut)
        Me.ToolStripStatusLabelIcon.Image = Me.RecentToolStripMenuItem.DropDown.Items(0).Image

        Me.PropertyGrid1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.StatusStrip1.ContextMenuStrip = Me.ContextMenuStrip1

        Me.LoadShortcutInHexBox(Me.CurrentShortcut.FullName)
    End Sub

    
    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Loads a shortcut file into the <see cref="Form1.HexBox1"/> control.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="filePath">
    ''' The shortcut file (.lnk) path.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub LoadShortcutInHexBox(filePath As String)
        
        If Me.currentFileByteProvider IsNot Nothing
            Me.currentFileByteProvider.Dispose()
            Me.currentFileByteProvider = Nothing
        End If

        Me.currentFileByteProvider = new Be.Windows.Forms.DynamicFileByteProvider(filePath, True)
        Me.HexBox1.ByteProvider = Nothing
        Me.HexBox1.ByteProvider = Me.currentFileByteProvider

    End Sub

    Private Sub AdjustPropertyGridSplitter()
        If Me.TabControl1.SelectedTab is Me.Tab_PropertyEditor 
            Try
                Me.PropertyGrid1.MoveSplitterTo(180)
            Catch
            End Try
        End If
    End Sub

#End Region

End Class

#End Region
