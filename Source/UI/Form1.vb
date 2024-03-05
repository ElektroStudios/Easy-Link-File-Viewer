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
Imports Manina.Windows.Forms

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
        Me.Text = $"{My.Application.Info.Title} {My.Application.Info.Version.Major}.{My.Application.Info.Version.Minor}"

        Me.ToolStripStatusLabelIcon.Text = ""
        Me.ToolStripStatusLabelFileName.Text = ""
        Me.AdjustCorrectTableLayoutPanelRowSizes()

        ' Load file from command-line arguments.
        If My.Application.CommandLineArgs.Any() Then
            Dim data As New DataObject(DataFormats.FileDrop, My.Application.CommandLineArgs.ToArray())
            Dim args As New DragEventArgs(data, Nothing, 0, 0, Nothing, DragDropEffects.Copy)
            Me.PropertyGrid1_DragDrop(Me, args)
        End If

        Me.AboutToolStripMenuItem.Alignment = ToolStripItemAlignment.Right
        Me.TabControl1.Pages.Remove(Me.Tab_HexViewer)
        JotUtil.StartTrackingMenuItems()
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
        'Dim minWidth As Integer = (From item As ToolStripMenuItem In Me.MenuStrip1.Items.Cast(Of ToolStripMenuItem)
        '                           Select item.Width + item.Padding.Horizontal).Sum()

        'Me.MinimumSize = New Size(minWidth, Me.Height)
        Me.MinimumSize = New Size(512, 512)
        If String.IsNullOrEmpty(Me.ToolStripComboBoxFontSize.Text) Then
            Me.ToolStripComboBoxFontSize.Text = "10"
        End If
        AddHandler Me.SettingsToolStripMenuItem.DropDown.Closing, AddressOf Me.SettingsToolStripDropDown_Closing
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
    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles NewToolStripMenuItem.Click, NewToolBarMenuItem.Click

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
                    MessageBox.Show(Me, "Can't delete existing link file:" & Environment.NewLine & Environment.NewLine & ex.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

                Dim newShortcut As ShortcutFileInfo = Nothing
                Try
                    newShortcut = New ShortcutFileInfo(dlg.FileName) With {.ViewMode = True}
                    newShortcut.Create()

                Catch ex As Exception
                    MessageBox.Show(Me, "Error creating link file:" & Environment.NewLine & Environment.NewLine & ex.Message,
                                    My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

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
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles OpenToolStripMenuItem.Click, OpenToolBarMenuItem.Click

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
    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles SaveToolStripMenuItem.Click, SaveToolBarMenuItem.Click

        Try
            Me.currentShortcut.Create()
            MessageBox.Show("Link file saved successfully.", My.Application.Info.Title,
                            MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(Me, "Error creating link file:" & Environment.NewLine & Environment.NewLine & ex.Message,
                             My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Me.PropertyGrid1.Refresh()
        Me.LoadShortcutInHexBox(Me.currentShortcut.FullName)
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
    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles SaveAsToolStripMenuItem.Click, SaveAsToolBarMenuItem.Click

        Using dlg As New SaveFileDialog()
            dlg.FileName = Me.currentShortcut.FullName
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
                    .Description = Me.currentShortcut.Description
                    .Hotkey = Me.currentShortcut.Hotkey
                    .Icon = Me.currentShortcut.Icon
                    .IconIndex = Me.currentShortcut.IconIndex
                    .Target = Me.currentShortcut.Target
                    .TargetArguments = Me.currentShortcut.TargetArguments
                    .WindowState = Me.currentShortcut.WindowState
                    .WorkingDirectory = Me.currentShortcut.WorkingDirectory

                    Try
                        .Create()
                        MessageBox.Show("Link file saved successfully.", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Information)

                    Catch ex As Exception
                        MessageBox.Show(Me, "Error creating link file:" & Environment.NewLine & Environment.NewLine & ex.Message,
                                    My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try


                    .Attributes = Me.currentShortcut.Attributes
                    .CreationTime = Me.currentShortcut.CreationTime
                    .LastAccessTime = Me.currentShortcut.LastAccessTime
                    .LastWriteTime = Me.currentShortcut.LastWriteTime
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
    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles CloseToolStripMenuItem.Click, CloseToolBarMenuItem.Click

        Me.HexBox1.ByteProvider = Nothing
        If Me.currentFileByteProvider IsNot Nothing Then
            Me.currentFileByteProvider.Dispose()
            Me.currentFileByteProvider = Nothing
        End If

        Me.PropertyGrid1.SelectedObject = Nothing
        Me.ToolStripStatusLabelFileName.Image = Nothing
        Me.ToolStripStatusLabelFileName.Text = ""

        Me.SaveToolStripMenuItem.Enabled = False
        Me.SaveAsToolStripMenuItem.Enabled = False
        Me.CloseToolStripMenuItem.Enabled = False

        Me.SaveToolBarMenuItem.Enabled = False
        Me.SaveAsToolBarMenuItem.Enabled = False
        Me.CloseToolBarMenuItem.Enabled = False

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
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

#End Region

#Region " Settings Menu "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripDropDown.Closing"/> event of the <see cref="Form1.SettingsToolStripMenuItem"/> control.
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
    Private Sub SettingsToolStripDropDown_Closing(sender As Object, e As ToolStripDropDownClosingEventArgs)
        e.Cancel = (e.CloseReason = ToolStripDropDownCloseReason.ItemClicked)
    End Sub

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
    Private Sub ToolStripComboBoxFontSize_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles ToolStripComboBoxFontSize.SelectedIndexChanged

        Dim fontSize As Single = CSng(DirectCast(sender, ToolStripComboBox).SelectedItem)
        Me.LoadFontSize(fontSize)

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.CheckedChanged"/> event of the <see cref="Form1.DefaultToolStripMenuItem"/> control.
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
    Private Sub DefaultToolStripMenuItem_CheckedChanged(sender As Object, e As EventArgs) _
        Handles DefaultToolStripMenuItem.CheckedChanged

        Dim item As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        If item.Checked Then
            Me.LoadVisualTheme(VisualStyle.Default)
        End If
    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.CheckedChanged"/> event of the <see cref="Form1.DarkToolStripMenuItem"/> control.
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
    Private Sub DarkToolStripMenuItem_CheckedChanged(sender As Object, e As EventArgs) _
        Handles DarkToolStripMenuItem.CheckedChanged

        Dim item As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        If item.Checked Then
            Me.LoadVisualTheme(VisualStyle.VisualStudioDark)
        End If
    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.Click"/> event of the
    ''' <see cref="Form1.DefaultToolStripMenuItem"/> and <see cref="Form1.DarkToolStripMenuItem"/> controls.
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
    Private Sub VisualStyles_ToolStripMenuItems_Click(sender As Object, e As EventArgs) _
        Handles DefaultToolStripMenuItem.Click, DarkToolStripMenuItem.Click

        ' Prevents from unchecking the checkbox when the item is checked.
        Dim item As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        If Not item.Checked Then
            item.Checked = True
        End If
    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.CheckedChanged"/> event of the <see cref="Form1.ShowFileMenuToolbarToolStripMenuItem"/> control.
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
    Private Sub ShowFileMenuToolbarToolStripMenuItem_CheckedChanged(sender As Object, e As EventArgs) _
        Handles ShowFileMenuToolbarToolStripMenuItem.CheckedChanged

        Me.AdjustCorrectTableLayoutPanelRowSizes()

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.CheckedChanged"/> event of the <see cref="Form1.ShowLinkEditorToolbarToolStripMenuItem"/> control.
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
    Private Sub ShowEditorToolbarToolStripMenuItem_CheckedChanged(sender As Object, e As EventArgs) _
        Handles ShowLinkEditorToolbarToolStripMenuItem.CheckedChanged

        Dim item As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        Me.PropertyGrid1.ToolbarVisible = item.Checked
    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.CheckedChanged"/> event of the <see cref="Form1.ShowRawTabToolStripMenuItem"/> control.
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
    Private Sub ShowRawTabToolStripMenuItem_CheckedChanged(sender As Object, e As EventArgs) _
        Handles ShowRawTabToolStripMenuItem.CheckedChanged

        Dim item As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        If Not item.Checked Then
            Me.TabControl1.SelectedTab = Me.Tab_PropertyEditor
            Me.TabControl1.Pages.Remove(Me.Tab_HexViewer)
        Else
            Me.TabControl1.Pages.Add(Me.Tab_HexViewer)
            Me.Tab_HexViewer.SetVisualStyle(If(Me.DefaultToolStripMenuItem.Checked, VisualStyle.Default, VisualStyle.VisualStudioDark))
            Me.HexBox1.SetVisualStyle(If(Me.DefaultToolStripMenuItem.Checked, VisualStyle.Default, VisualStyle.VisualStudioDark))
            If Me.currentShortcut IsNot Nothing Then
                Me.LoadShortcutInHexBox(Me.currentShortcut.FullName)
            End If
        End If

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.CheckedChanged"/> event of the <see cref="Form1.RememberWindowSizeAndPosToolStripMenuItem"/> control.
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
    Private Sub RememberWindowSizeAndPosToolStripMenuItem_CheckedChanged(sender As Object, e As EventArgs) _
    Handles RememberWindowSizeAndPosToolStripMenuItem.CheckedChanged

        Dim item As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        If item.Checked Then
            JotUtil.StartTrackingForm()
        Else
            JotUtil.StopTrackingForm()
        End If

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.CheckedChanged"/> event of the <see cref="Form1.AddProgramShortcutToExplorersContextmenuToolStripMenuItem"/> control.
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
    Private Sub AddProgramShortcutToExplorersContextmenuToolStripMenuItem_CheckedChanged(sender As Object, e As EventArgs) _
        Handles AddProgramShortcutToExplorersContextmenuToolStripMenuItem.CheckedChanged

        Dim executablePath As String = Application.ExecutablePath()
        Dim fileType As String = "lnkfile"
        Dim keyName As String = "OpenInEasyLinkFileViewer"
        Dim text As String = $"Open in Easy Link File Viewer"
        Dim position As String = "middle"
        Dim icon As String = $"{executablePath},0"
        Dim command As String = $"""{executablePath}"" ""%1"""

        Dim item As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        If item.Checked Then
            Try
                RegistryUtil.CreateFileTypeRegistryMenuEntry(fileType, keyName, text, position, icon, command)
            Catch ex As Exception
                MessageBox.Show(Me, "Error trying to create file type registry menu entry:" & Environment.NewLine & Environment.NewLine & ex.Message,
                                My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            Try
                RegistryUtil.DeleteFileTypeRegistryMenuEntry(fileType, keyName, throwOnMissingsubKey:=False)
            Catch ex As Exception
                MessageBox.Show(Me, "Error trying to delete file type registry menu entry:" & Environment.NewLine & Environment.NewLine & ex.Message,
                                My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)

            End Try
        End If

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="ToolStripMenuItem.CheckedChanged"/> event of the <see cref="Form1.DefaultToolStripMenuItem"/> control.
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
    Private Sub DisableRecentFilesListToolStripMenuItem_CheckedChanged(sender As Object, e As EventArgs) _
        Handles HideRecentFilesListToolStripMenuItem.CheckedChanged
        ' MsgBox(1)
        Dim item As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)

        If item.Checked Then
            Me.FileToolStripMenuItem.DropDownItems.Remove(Me.RecentToolStripMenuItem)
            Me.FileToolStripMenuItem.DropDownItems.Remove(Me.ToolStripSeparator1)
        Else
            Me.FileToolStripMenuItem.DropDownItems.Insert(5, Me.RecentToolStripMenuItem)
            Me.FileToolStripMenuItem.DropDownItems.Insert(5, Me.ToolStripSeparator1)
        End If

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

        Me.OpenShortcutMenuItem.Enabled = Me.currentShortcut.Exists
        Me.ViewShortcutMenuItem.Enabled = Me.OpenShortcutMenuItem.Enabled

        Me.OpenTargetMenuItem.Enabled = File.Exists(Me.currentShortcut.Target) OrElse Directory.Exists(Me.currentShortcut.Target)
        Me.OpenTargetWithArgsMenuItem.Enabled = File.Exists(Me.currentShortcut.Target) AndAlso Not String.IsNullOrWhiteSpace(Me.currentShortcut.TargetArguments)
        Me.ViewTargetMenuItem.Enabled = Me.OpenTargetMenuItem.Enabled

        Me.ViewWorkingDirectoryMenuItem.Enabled = Directory.Exists(Me.currentShortcut.WorkingDirectory)
        Me.ViewIconMenuItem.Enabled = File.Exists(Me.currentShortcut.Icon)

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
            Process.Start(Me.currentShortcut.FullName)

        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)

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

        Dim target As String = Me.currentShortcut.Target

        If File.Exists(target) Then
            Try
                Process.Start(target)
                Exit Sub

            Catch ex As Exception
                MessageBox.Show(Me, ex.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)

            End Try
        End If

        If Directory.Exists(target) Then
            Try
                FileUtil.InternalOpenInExplorer(target)
                Exit Sub

            Catch ex As Exception
                MessageBox.Show(Me, ex.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)

            End Try
        End If

        MessageBox.Show(Me, "Can't find the shortcut target.", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)

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
            Process.Start(Me.currentShortcut.Target, Me.currentShortcut.TargetArguments)
            Exit Sub

        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)

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
            FileUtil.InternalOpenInExplorer(Me.currentShortcut.FullName)

        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)

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
            FileUtil.InternalOpenInExplorer(Me.currentShortcut.Target)

        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)

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
            FileUtil.InternalOpenInExplorer(Me.currentShortcut.WorkingDirectory)

        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)

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
            FileUtil.InternalOpenInExplorer(Me.currentShortcut.Icon)

        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)

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
    Private Sub LoadFontSize(fontSize As Single)
        Me.MenuStrip1.Font = New Font(Me.MenuStrip1.Font.FontFamily, fontSize, Me.MenuStrip1.Font.Style)
        Me.StatusStrip1.Font = New Font(Me.StatusStrip1.Font.FontFamily, fontSize, Me.StatusStrip1.Font.Style)
        Me.ToolStripStatusLabelIcon.Font = New Font(Me.ToolStripStatusLabelIcon.Font.FontFamily, fontSize, Me.ToolStripStatusLabelIcon.Font.Style)
        Me.ToolStripStatusLabelFileName.Font = New Font(Me.ToolStripStatusLabelFileName.Font.FontFamily, fontSize, Me.ToolStripStatusLabelFileName.Font.Style)
        Me.PropertyGrid1.Font = New Font(Me.PropertyGrid1.Font.FontFamily, fontSize, Me.PropertyGrid1.Font.Style)

        For Each tab As Tab In Me.TabControl1.Tabs
            tab.Font = New Font(tab.Font.FontFamily, fontSize, tab.Font.Style)
        Next

        Me.MenuStrip_ToolBar.Font = New Font(Me.MenuStrip_ToolBar.Font.FontFamily, fontSize, Me.MenuStrip_ToolBar.Font.Style)
        Me.ToolStripComboBoxFontSize.Font = New Font(Me.ToolStripComboBoxFontSize.Font.FontFamily, fontSize, Me.ToolStripComboBoxFontSize.Font.Style)
        My.Forms.AboutBox1.Font = New Font(My.Forms.AboutBox1.Font.FontFamily, fontSize, My.Forms.AboutBox1.Font.Style)
        My.Forms.AboutBox1.LinkLabel1.Font = New Font(My.Forms.AboutBox1.LinkLabel1.Font.FontFamily, fontSize, My.Forms.AboutBox1.LinkLabel1.Font.Style)

        Me.HexBox1.Font = New Font(Me.HexBox1.Font.FontFamily, fontSize, Me.HexBox1.Font.Style)
        Me.AdjustCorrectTableLayoutPanelRowSizes()
    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Loads the saved visual theme for the user-interface.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub LoadVisualTheme(visualStyle As VisualStyle)
        Select Case visualStyle

            Case VisualStyle.Default
                Me.SetVisualStyle(VisualStyle.Default, True)
                My.Forms.AboutBox1.SetVisualStyle(VisualStyle.Default, True)

                Me.DefaultToolStripMenuItem.Checked = True
                Me.DarkToolStripMenuItem.Checked = False

            Case VisualStyle.VisualStudioDark
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
        Try
            Me.currentShortcut = New ShortcutFileInfo(filePath) With {.ViewMode = True}
        Catch ex As Exception
            MessageBox.Show(Me, "Error trying to read link data:" & Environment.NewLine & Environment.NewLine & ex.Message,
                            My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
        If Not Me.currentShortcut.Exists Then
            MessageBox.Show(Me, $"The lnk file does not exist: {Me.currentShortcut.FullName}", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If Me.currentShortcut.IsWindowsInstallerShortcut Then
            MessageBox.Show(Me, "This link points to a Windows Installer product." & Environment.NewLine & Environment.NewLine &
                                "Support to read/write this link is experimental and saving changes to this file could corrupt the link. Please save any changes to a new link file.",
                            My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

        Me.AdjustPropertyGridSplitter()
        Me.PropertyGrid1.SelectedObject = Me.currentShortcut
        Me.ToolStripStatusLabelFileName.Text = filePath

        Me.RecentToolStripMenuItem.Enabled = True
        Me.SaveToolStripMenuItem.Enabled = True
        Me.SaveAsToolStripMenuItem.Enabled = True
        Me.CloseToolStripMenuItem.Enabled = True

        Me.SaveToolBarMenuItem.Enabled = True
        Me.SaveAsToolBarMenuItem.Enabled = True
        Me.CloseToolBarMenuItem.Enabled = True

        Me.UpdateMruItems(Me.currentShortcut)
        Me.ToolStripStatusLabelIcon.Image = Me.RecentToolStripMenuItem.DropDown.Items(0).Image

        Me.PropertyGrid1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.StatusStrip1.ContextMenuStrip = Me.ContextMenuStrip1

        Me.LoadShortcutInHexBox(Me.currentShortcut.FullName)
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

        If Not Me.TabControl1.Pages.Contains(Me.Tab_HexViewer) Then
            Exit Sub
        End If

        If Me.currentFileByteProvider IsNot Nothing Then
            Me.currentFileByteProvider.Dispose()
            Me.currentFileByteProvider = Nothing
        End If

        Try
            Me.currentFileByteProvider = New Be.Windows.Forms.DynamicFileByteProvider(filePath, True)
            Me.HexBox1.ByteProvider = Nothing
            Me.HexBox1.ByteProvider = Me.currentFileByteProvider
        Catch ex As Exception
            MessageBox.Show(Me, "Error trying to read link RAW data:" & Environment.NewLine & Environment.NewLine & ex.Message,
                            My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub AdjustPropertyGridSplitter()
        If Me.TabControl1.SelectedTab Is Me.Tab_PropertyEditor Then
            Try
                Me.PropertyGrid1.MoveSplitterTo(180)
            Catch
            End Try
        End If
    End Sub

    Private Sub AdjustCorrectTableLayoutPanelRowSizes()

        Dim item As ToolStripMenuItem = Me.ShowFileMenuToolbarToolStripMenuItem
        Dim makeToolBarVisible As Boolean = item.Checked
        Dim table As TableLayoutPanel = Me.TableLayoutPanel1
        Dim rowIndex As Integer = table.GetRow(Me.MenuStrip_ToolBar)
        Dim isTableRowsInitialized As Boolean = (rowIndex <> -1)

        If Not makeToolBarVisible Then
            If isTableRowsInitialized Then
                table.RowStyles(rowIndex).Height = 0
            End If
            Me.MenuStrip_ToolBar.Visible = False

        Else
            Me.MenuStrip_ToolBar.Visible = True
            If isTableRowsInitialized Then
                table.RowStyles(rowIndex).Height = 80
                table.RowStyles(rowIndex).Height = Me.MenuStrip_ToolBar.Height
            End If
        End If

    End Sub

#End Region

End Class

#End Region
