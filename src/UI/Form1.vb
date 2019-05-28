' ***********************************************************************
' Author   : ElektroStudios
' Modified : 27-May-2019
' ***********************************************************************

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports DevCase.Core.Extensions
Imports DevCase.Core.IO
Imports DevCase.Core.Imaging.Tools
Imports System.IO

#End Region

#Region " Form1 "

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' The main application Form.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
Friend NotInheritable Class Form1 : Inherits Form

#Region " Constructors "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Initializes a new instance of the <see cref="Form1"/> class.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Public Sub New()
        MyClass.InitializeComponent()
    End Sub

#End Region

#Region " Event Handlers "

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
        Me.LoadVisualTheme()
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
        Dim minWidth As Integer = (From item As ToolStripMenuItem In MenuStrip1.Items.Cast(Of ToolStripMenuItem)
                                   Select item.Width + item.Padding.Horizontal).Sum()

        Me.MinimumSize = New Size(minWidth, Me.Size.Height)
    End Sub

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

        Using ofd As New OpenFileDialog
            ofd.DefaultExt = "lnk"
            ofd.DereferenceLinks = False
            ofd.Filter = "Shortcut files (*.lnk)|*.lnk"
            ofd.FilterIndex = 1
            ofd.Multiselect = False
            ofd.RestoreDirectory = True
            ofd.SupportMultiDottedExtensions = True
            ofd.Title = "Select a shortcut file to open..."

            If ofd.ShowDialog = DialogResult.OK Then
                Me.LoadShortcutInPropertyGrid(ofd.FileName)
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

        Dim shortcut As ShortcutFileInfo = DirectCast(Me.PropertyGrid1.SelectedObject, ShortcutFileInfo)
        shortcut.Create()

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

        Dim shortcut As ShortcutFileInfo = DirectCast(Me.PropertyGrid1.SelectedObject, ShortcutFileInfo)

        Using ofd As New SaveFileDialog
            ofd.FileName = shortcut.FullName
            ofd.DefaultExt = "lnk"
            ofd.DereferenceLinks = False
            ofd.Filter = "Shortcut files (*.lnk)|*.lnk"
            ofd.FilterIndex = 1
            ofd.RestoreDirectory = True
            ofd.SupportMultiDottedExtensions = True
            ofd.Title = "Select a destination to save the shortcut file..."

            If ofd.ShowDialog() = DialogResult.OK Then

                Dim dstShortcut As New ShortcutFileInfo(ofd.FileName) With {.ViewMode = True}
                With dstShortcut
                    .Description = shortcut.Description
                    .Hotkey = shortcut.Hotkey
                    .Icon = shortcut.Icon
                    .IconIndex = shortcut.IconIndex
                    .Target = shortcut.Target
                    .TargetArguments = shortcut.TargetArguments
                    .WindowState = shortcut.WindowState
                    .WorkingDirectory = shortcut.WorkingDirectory

                    .Create()

                    .Attributes = shortcut.Attributes
                    .CreationTime = shortcut.CreationTime
                    .LastAccessTime = shortcut.LastAccessTime
                    .LastWriteTime = shortcut.LastWriteTime
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
        Me.PropertyGrid1.SelectedObject = Nothing
        Me.ToolStripStatusLabel1.Text = ""

        Me.SaveToolStripMenuItem.Enabled = False
        Me.SaveAsToolStripMenuItem.Enabled = False
        Me.CloseToolStripMenuItem.Enabled = False
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

#Region " Private Methods "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Loads the saved visual theme for the user-interface.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub LoadVisualTheme()
        Select Case My.Settings.VisualThemeIndex

            Case 0 ' Default theme
                Me.SetThemeDefault(True)
                My.Forms.AboutBox1.SetThemeDefault(True)

                Me.DefaultToolStripMenuItem.Checked = True
                Me.DarkToolStripMenuItem.Checked = False

            Case 1 ' Dark theme
                Me.SetThemeVisualStudioDark(True)
                My.Forms.AboutBox1.SetThemeVisualStudioDark(True)

                Me.DefaultToolStripMenuItem.Checked = False
                Me.DarkToolStripMenuItem.Checked = True

            Case Else

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
    Private Sub UpdateMruItems(ByVal shortcut As ShortcutFileInfo)

        ' Search and remove existing MRU item.
        For Each item As ToolStripMenuItem In Me.RecentToolStripMenuItem.DropDown.Items
            If CStr(item.Tag).Equals(shortcut.FullName, StringComparison.OrdinalIgnoreCase) Then
                item.Image?.Dispose()
                item.Dispose()
                Exit For
            End If
        Next

        ' Keep maximum capacity.
        Const MaxMruCapacity As Integer = 10
        If (Me.RecentToolStripMenuItem.DropDown.Items.Count = MaxMruCapacity) Then
            Me.RecentToolStripMenuItem.DropDown.Items(MaxMruCapacity - 1).Dispose()
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
    Private Sub LoadShortcutInPropertyGrid(ByVal filePath As String)

        Dim shortcut As New ShortcutFileInfo(filePath) With {.ViewMode = True}
        If Not shortcut.Exists Then
            MessageBox.Show(Me, $"The lnk file does not exist: {shortcut.FullName}", Me.Name, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Me.PropertyGrid1.MoveSplitterTo(160)
        Me.PropertyGrid1.SelectedObject = shortcut
        Me.ToolStripStatusLabel1.Text = filePath

        Me.RecentToolStripMenuItem.Enabled = True
        Me.SaveToolStripMenuItem.Enabled = True
        Me.SaveAsToolStripMenuItem.Enabled = True
        Me.CloseToolStripMenuItem.Enabled = True

        Me.UpdateMruItems(shortcut)

    End Sub

#End Region

End Class

#End Region