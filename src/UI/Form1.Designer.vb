<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Friend Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RecentToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.OpenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveAsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CloseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SettingsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FontSizeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripComboBoxFontSize = New System.Windows.Forms.ToolStripComboBox()
        Me.VisualThemeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DefaultToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DarkToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PropertyGrid1 = New System.Windows.Forms.PropertyGrid()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.OpenShortcutMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenTargetMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenTargetWithArgsMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewShortcutMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewTargetMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewWorkingDirectoryMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewIconMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabelIcon = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabelFileName = New System.Windows.Forms.ToolStripStatusLabel()
        Me.MenuStrip1.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Font = New System.Drawing.Font("Segoe UI", 10.0!)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.SettingsToolStripMenuItem, Me.AboutToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(684, 27)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewToolStripMenuItem, Me.OpenToolStripMenuItem, Me.SaveToolStripMenuItem, Me.SaveAsToolStripMenuItem, Me.CloseToolStripMenuItem, Me.ToolStripSeparator1, Me.RecentToolStripMenuItem, Me.ToolStripSeparator2, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Image = Global.My.Resources.Resources.FSFile_16x
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(57, 23)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'NewToolStripMenuItem
        '
        Me.NewToolStripMenuItem.Image = Global.My.Resources.Resources.NewFile_16x
        Me.NewToolStripMenuItem.Name = "NewToolStripMenuItem"
        Me.NewToolStripMenuItem.Size = New System.Drawing.Size(180, 24)
        Me.NewToolStripMenuItem.Text = "New"
        '
        'RecentToolStripMenuItem
        '
        Me.RecentToolStripMenuItem.Enabled = False
        Me.RecentToolStripMenuItem.Image = Global.My.Resources.Resources.FileGroup_16x
        Me.RecentToolStripMenuItem.Name = "RecentToolStripMenuItem"
        Me.RecentToolStripMenuItem.Size = New System.Drawing.Size(180, 24)
        Me.RecentToolStripMenuItem.Text = "Recent..."
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(177, 6)
        '
        'OpenToolStripMenuItem
        '
        Me.OpenToolStripMenuItem.Image = Global.My.Resources.Resources.OpenFile_16x
        Me.OpenToolStripMenuItem.Name = "OpenToolStripMenuItem"
        Me.OpenToolStripMenuItem.Size = New System.Drawing.Size(180, 24)
        Me.OpenToolStripMenuItem.Text = "Open"
        '
        'SaveToolStripMenuItem
        '
        Me.SaveToolStripMenuItem.Enabled = False
        Me.SaveToolStripMenuItem.Image = Global.My.Resources.Resources.Save_16x
        Me.SaveToolStripMenuItem.Name = "SaveToolStripMenuItem"
        Me.SaveToolStripMenuItem.Size = New System.Drawing.Size(180, 24)
        Me.SaveToolStripMenuItem.Text = "Save"
        '
        'SaveAsToolStripMenuItem
        '
        Me.SaveAsToolStripMenuItem.Enabled = False
        Me.SaveAsToolStripMenuItem.Image = Global.My.Resources.Resources.Save_16x
        Me.SaveAsToolStripMenuItem.Name = "SaveAsToolStripMenuItem"
        Me.SaveAsToolStripMenuItem.Size = New System.Drawing.Size(180, 24)
        Me.SaveAsToolStripMenuItem.Text = "Save As..."
        '
        'CloseToolStripMenuItem
        '
        Me.CloseToolStripMenuItem.Enabled = False
        Me.CloseToolStripMenuItem.Image = Global.My.Resources.Resources.FileExclude_16x
        Me.CloseToolStripMenuItem.Name = "CloseToolStripMenuItem"
        Me.CloseToolStripMenuItem.Size = New System.Drawing.Size(180, 24)
        Me.CloseToolStripMenuItem.Text = "Close"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(177, 6)
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Image = Global.My.Resources.Resources.Close_red_16x
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(180, 24)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'SettingsToolStripMenuItem
        '
        Me.SettingsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FontSizeToolStripMenuItem, Me.VisualThemeToolStripMenuItem})
        Me.SettingsToolStripMenuItem.Image = Global.My.Resources.Resources.Settings_16x
        Me.SettingsToolStripMenuItem.Name = "SettingsToolStripMenuItem"
        Me.SettingsToolStripMenuItem.Size = New System.Drawing.Size(86, 23)
        Me.SettingsToolStripMenuItem.Text = "Settings"
        '
        'FontSizeToolStripMenuItem
        '
        Me.FontSizeToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripComboBoxFontSize})
        Me.FontSizeToolStripMenuItem.Image = Global.My.Resources.Resources.FontSize_16x
        Me.FontSizeToolStripMenuItem.Name = "FontSizeToolStripMenuItem"
        Me.FontSizeToolStripMenuItem.Size = New System.Drawing.Size(159, 24)
        Me.FontSizeToolStripMenuItem.Text = "Font Size"
        '
        'ToolStripComboBoxFontSize
        '
        Me.ToolStripComboBoxFontSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ToolStripComboBoxFontSize.Items.AddRange(New Object() {"8", "9", "10", "11", "12"})
        Me.ToolStripComboBoxFontSize.Name = "ToolStripComboBoxFontSize"
        Me.ToolStripComboBoxFontSize.Size = New System.Drawing.Size(121, 23)
        '
        'VisualThemeToolStripMenuItem
        '
        Me.VisualThemeToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DefaultToolStripMenuItem, Me.DarkToolStripMenuItem})
        Me.VisualThemeToolStripMenuItem.Image = Global.My.Resources.Resources.ColorPalette_16x
        Me.VisualThemeToolStripMenuItem.Name = "VisualThemeToolStripMenuItem"
        Me.VisualThemeToolStripMenuItem.Size = New System.Drawing.Size(159, 24)
        Me.VisualThemeToolStripMenuItem.Text = "Visual Theme"
        '
        'DefaultToolStripMenuItem
        '
        Me.DefaultToolStripMenuItem.Image = Global.My.Resources.Resources.ColorWheel_16x
        Me.DefaultToolStripMenuItem.Name = "DefaultToolStripMenuItem"
        Me.DefaultToolStripMenuItem.Size = New System.Drawing.Size(122, 24)
        Me.DefaultToolStripMenuItem.Text = "Default"
        '
        'DarkToolStripMenuItem
        '
        Me.DarkToolStripMenuItem.Image = Global.My.Resources.Resources.DarkTheme_16x
        Me.DarkToolStripMenuItem.Name = "DarkToolStripMenuItem"
        Me.DarkToolStripMenuItem.Size = New System.Drawing.Size(122, 24)
        Me.DarkToolStripMenuItem.Text = "Dark"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Image = Global.My.Resources.Resources.Question_16x
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(84, 23)
        Me.AboutToolStripMenuItem.Text = "About..."
        '
        'PropertyGrid1
        '
        Me.PropertyGrid1.AllowDrop = True
        Me.PropertyGrid1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PropertyGrid1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.PropertyGrid1.Location = New System.Drawing.Point(0, 27)
        Me.PropertyGrid1.Name = "PropertyGrid1"
        Me.PropertyGrid1.Size = New System.Drawing.Size(684, 360)
        Me.PropertyGrid1.TabIndex = 1
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OpenShortcutMenuItem, Me.OpenTargetMenuItem, Me.OpenTargetWithArgsMenuItem, Me.ViewShortcutMenuItem, Me.ViewTargetMenuItem, Me.ViewWorkingDirectoryMenuItem, Me.ViewIconMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(257, 158)
        '
        'OpenShortcutMenuItem
        '
        Me.OpenShortcutMenuItem.Image = Global.My.Resources.Resources.Open_16x
        Me.OpenShortcutMenuItem.Name = "OpenShortcutMenuItem"
        Me.OpenShortcutMenuItem.Size = New System.Drawing.Size(256, 22)
        Me.OpenShortcutMenuItem.Text = "Open Shortcut"
        '
        'OpenTargetMenuItem
        '
        Me.OpenTargetMenuItem.Image = Global.My.Resources.Resources.Open_16x
        Me.OpenTargetMenuItem.Name = "OpenTargetMenuItem"
        Me.OpenTargetMenuItem.Size = New System.Drawing.Size(256, 22)
        Me.OpenTargetMenuItem.Text = "Open Target"
        '
        'OpenTargetWithArgsMenuItem
        '
        Me.OpenTargetWithArgsMenuItem.Image = Global.My.Resources.Resources.Open_16x
        Me.OpenTargetWithArgsMenuItem.Name = "OpenTargetWithArgsMenuItem"
        Me.OpenTargetWithArgsMenuItem.Size = New System.Drawing.Size(256, 22)
        Me.OpenTargetWithArgsMenuItem.Text = "Open Target with Arguments"
        '
        'ViewShortcutMenuItem
        '
        Me.ViewShortcutMenuItem.Image = Global.My.Resources.Resources.FolderOpen_16x
        Me.ViewShortcutMenuItem.Name = "ViewShortcutMenuItem"
        Me.ViewShortcutMenuItem.Size = New System.Drawing.Size(256, 22)
        Me.ViewShortcutMenuItem.Text = "View Shortcut in Explorer"
        '
        'ViewTargetMenuItem
        '
        Me.ViewTargetMenuItem.Image = Global.My.Resources.Resources.FolderOpen_16x
        Me.ViewTargetMenuItem.Name = "ViewTargetMenuItem"
        Me.ViewTargetMenuItem.Size = New System.Drawing.Size(256, 22)
        Me.ViewTargetMenuItem.Text = "View Target in Explorer"
        '
        'ViewWorkingDirectoryMenuItem
        '
        Me.ViewWorkingDirectoryMenuItem.Image = Global.My.Resources.Resources.FolderOpen_16x
        Me.ViewWorkingDirectoryMenuItem.Name = "ViewWorkingDirectoryMenuItem"
        Me.ViewWorkingDirectoryMenuItem.Size = New System.Drawing.Size(256, 22)
        Me.ViewWorkingDirectoryMenuItem.Text = "View Working Directory in Explorer"
        '
        'ViewIconMenuItem
        '
        Me.ViewIconMenuItem.Image = Global.My.Resources.Resources.FolderOpen_16x
        Me.ViewIconMenuItem.Name = "ViewIconMenuItem"
        Me.ViewIconMenuItem.Size = New System.Drawing.Size(256, 22)
        Me.ViewIconMenuItem.Text = "View Icon in Explorer"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.BackColor = System.Drawing.SystemColors.Control
        Me.StatusStrip1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabelIcon, Me.ToolStripStatusLabelFileName})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 387)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(684, 24)
        Me.StatusStrip1.TabIndex = 2
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabelIcon
        '
        Me.ToolStripStatusLabelIcon.Font = New System.Drawing.Font("Segoe UI", 10.0!)
        Me.ToolStripStatusLabelIcon.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolStripStatusLabelIcon.Name = "ToolStripStatusLabelIcon"
        Me.ToolStripStatusLabelIcon.Size = New System.Drawing.Size(42, 19)
        Me.ToolStripStatusLabelIcon.Text = "{icon}"
        Me.ToolStripStatusLabelIcon.TextImageRelation = System.Windows.Forms.TextImageRelation.Overlay
        '
        'ToolStripStatusLabelFileName
        '
        Me.ToolStripStatusLabelFileName.AutoToolTip = True
        Me.ToolStripStatusLabelFileName.Font = New System.Drawing.Font("Segoe UI", 10.0!)
        Me.ToolStripStatusLabelFileName.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolStripStatusLabelFileName.LinkBehavior = System.Windows.Forms.LinkBehavior.NeverUnderline
        Me.ToolStripStatusLabelFileName.Name = "ToolStripStatusLabelFileName"
        Me.ToolStripStatusLabelFileName.Size = New System.Drawing.Size(596, 19)
        Me.ToolStripStatusLabelFileName.Spring = True
        Me.ToolStripStatusLabelFileName.Text = "{fullpath}     "
        Me.ToolStripStatusLabelFileName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(684, 411)
        Me.Controls.Add(Me.PropertyGrid1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Easy Link File Viewer"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents OpenToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SaveToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SaveAsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CloseToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SettingsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents VisualThemeToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents DefaultToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents DarkToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents PropertyGrid1 As PropertyGrid
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents ToolStripStatusLabelFileName As ToolStripStatusLabel
    Friend WithEvents RecentToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents OpenShortcutMenuItem As ToolStripMenuItem
    Friend WithEvents ViewShortcutMenuItem As ToolStripMenuItem
    Friend WithEvents ViewTargetMenuItem As ToolStripMenuItem
    Friend WithEvents ViewWorkingDirectoryMenuItem As ToolStripMenuItem
    Friend WithEvents OpenTargetMenuItem As ToolStripMenuItem
    Friend WithEvents OpenTargetWithArgsMenuItem As ToolStripMenuItem
    Friend WithEvents ViewIconMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents FontSizeToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripComboBoxFontSize As ToolStripComboBox
    Friend WithEvents ToolStripStatusLabelIcon As ToolStripStatusLabel
    Friend WithEvents NewToolStripMenuItem As ToolStripMenuItem
End Class
