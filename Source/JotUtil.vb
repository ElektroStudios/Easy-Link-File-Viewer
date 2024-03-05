#Region " Option Statements "

Option Explicit On
Option Strict On
Option Infer Off

#End Region

#Region " Imports "

Imports Microsoft.VisualBasic.ApplicationServices

Imports Jot
Imports Jot.Configuration
Imports Jot.Storage

#End Region
''' <summary>
''' Provides JOT related utilities.
''' </summary>
Friend Module JotUtil

#Region " Fields "

    ''' <summary>
    ''' Directory path where to save the Jot cache data.
    ''' </summary>
    Friend JotCachePath As String = $"{My.Application.Info.DirectoryPath}\cache\jot"

    ''' <summary>
    ''' The object responsible for tracking the specified properties of the specified target objects
    ''' in <see cref="JotUtil.formTrackingConfig"/> and <see cref="JotUtil.menuItemTrackingConfig"/>.
    ''' <para></para>
    ''' Tracking means persisting the values of the specified object properties,
    ''' and restoring this data when appropriate.
    ''' </summary>
    Private jotTracker As Tracker

    ''' <summary>
    ''' The object that determines how a target <see cref="Form"/> will be tracked. 
    ''' <para></para>
    ''' This includes list of properties to track, persist triggers and id getter.
    ''' </summary>
    Private formTrackingConfig As TrackingConfiguration(Of Form1)

    ''' <summary>
    ''' The object that determines how a target <see cref="PropertyGrid"/> will be tracked. 
    ''' <para></para>
    ''' This includes list of properties to track, persist triggers and id getter.
    ''' </summary>
    Private propertyGridTrackingConfig As TrackingConfiguration(Of PropertyGrid)

    ''' <summary>
    ''' The object that determines how a target <see cref="ToolStripMenuItem"/> will be tracked. 
    ''' <para></para>
    ''' This includes list of properties to track, persist triggers and id getter.
    ''' </summary>
    Private menuItemTrackingConfig As TrackingConfiguration(Of ToolStripMenuItem)

    ''' <summary>
    ''' The object that determines how a target <see cref="ToolStripComboBox"/> will be tracked. 
    ''' <para></para>
    ''' This includes list of properties to track, persist triggers and id getter.
    ''' </summary>
    Private toolStripComboBoxTrackingConfig As TrackingConfiguration(Of ToolStripComboBox)

#End Region

#Region " Public Methods "

    ''' <summary>
    ''' Initializes JOT environment.
    ''' <para></para>
    ''' This method should be called on your main application thread (typically the UI thread)
    ''' and typically from the <see cref="WindowsFormsApplicationBase.Startup"/> event,
    ''' before the <see cref="MainForm"/> form is loaded.
    ''' </summary>
    <DebuggerStepperBoundary>
    Friend Sub InitializeJot()

        JotUtil.jotTracker = New Tracker()
        DirectCast(JotUtil.jotTracker.Store, JsonFileStore).FolderPath = JotUtil.JotCachePath

        JotUtil.formTrackingConfig = JotUtil.jotTracker.Configure(Of Form1)
        With JotUtil.formTrackingConfig
            .Id(Function(f As Form1) f.Name, SystemInformation.VirtualScreen.Size)
            .Properties(Function(f As Form1) New With {f.Top, f.Width, f.Height, f.Left, f.WindowState})
            .PersistOn(NameOf(Form1.Move), NameOf(Form1.Resize), NameOf(Form1.FormClosing))
            .WhenPersistingProperty(
                Sub(f As Form1, pod As PropertyOperationData)
                    pod.Cancel = f.WindowState <> FormWindowState.Normal AndAlso
                               (pod.Property = NameOf(Form1.Height) OrElse
                                pod.Property = NameOf(Form1.Width) OrElse
                                pod.Property = NameOf(Form1.Top) OrElse
                                pod.Property = NameOf(Form1.Left))
                End Sub)
            .StopTrackingOn(NameOf(Form1.FormClosing))
            .CanPersist(Function(f As Form1) f.RememberWindowSizeAndPosToolStripMenuItem.Checked)
        End With

        JotUtil.propertyGridTrackingConfig = JotUtil.jotTracker.Configure(Of PropertyGrid)
        With JotUtil.propertyGridTrackingConfig
            .Id(Function(pg As PropertyGrid) pg.Name)
            .Properties(Function(pg As PropertyGrid) New With {pg.PropertySort})
            .PersistOn(NameOf(PropertyGrid.PropertySortChanged))
            .StopTrackingOn(NameOf(PropertyGrid.Disposed))
        End With

        JotUtil.menuItemTrackingConfig = JotUtil.jotTracker.Configure(Of ToolStripMenuItem)
        With JotUtil.menuItemTrackingConfig
            .Id(Function(i As ToolStripMenuItem) i.Name)
            .Properties(Function(i As ToolStripMenuItem) New With {i.Checked})
            .PersistOn(NameOf(ToolStripMenuItem.CheckedChanged))
            .StopTrackingOn(NameOf(ToolStripMenuItem.Disposed))
        End With

        JotUtil.toolStripComboBoxTrackingConfig = JotUtil.jotTracker.Configure(Of ToolStripComboBox)
        With JotUtil.toolStripComboBoxTrackingConfig
            .Id(Function(i As ToolStripComboBox) i.Name)
            .Properties(Function(i As ToolStripComboBox) New With {i.SelectedIndex})
            .PersistOn(NameOf(ToolStripComboBox.SelectedIndexChanged))
            .StopTrackingOn(NameOf(ToolStripComboBox.Disposed))
        End With

    End Sub

    ''' <summary>
    ''' Starts tracking the size and location of the <see cref="MainForm"/> form.
    ''' <para></para>
    ''' This method should be called in <see cref="Form1.Load"/> event handler.
    ''' </summary>
    <DebuggerStepThrough>
    Friend Sub StartTrackingForm()
        JotUtil.formTrackingConfig.Track(My.Forms.Form1)
        JotUtil.propertyGridTrackingConfig.Track(My.Forms.Form1.PropertyGrid1)
    End Sub

    ''' <summary>
    ''' Starts tracking the size and location of the <see cref="MainForm"/> form.
    ''' <para></para>
    ''' This method should be called in <see cref="Form1.Load"/> event handler.
    ''' </summary>
    <DebuggerStepThrough>
    Friend Sub StopTrackingForm()
        JotUtil.jotTracker.Forget(My.Forms.Form1)
        JotUtil.jotTracker.ApplyDefaults(My.Forms.Form1)
        Dim storeId As String = JotUtil.formTrackingConfig.GetStoreId(My.Forms.Form1)
        DirectCast(JotUtil.jotTracker.Store, JsonFileStore).ClearData(storeId)
    End Sub

    ''' <summary>
    ''' Starts tracking the <see cref="ToolStripMenuItem"/> checked states.
    ''' <para></para>
    ''' This method should be called in <see cref="Form1.Shown"/> event handler, 
    ''' after the ChromiumWebBrowser is initialized.
    ''' </summary>
    <DebuggerStepThrough>
    Friend Sub StartTrackingMenuItems()
        JotUtil.menuItemTrackingConfig.Track(My.Forms.Form1.RememberWindowSizeAndPosToolStripMenuItem)

        JotUtil.menuItemTrackingConfig.Track(My.Forms.Form1.DefaultToolStripMenuItem)
        JotUtil.menuItemTrackingConfig.Track(My.Forms.Form1.DarkToolStripMenuItem)
        JotUtil.menuItemTrackingConfig.Track(My.Forms.Form1.ShowFileMenuToolbarToolStripMenuItem)
        JotUtil.menuItemTrackingConfig.Track(My.Forms.Form1.ShowLinkEditorToolbarToolStripMenuItem)
        JotUtil.menuItemTrackingConfig.Track(My.Forms.Form1.ShowRawTabToolStripMenuItem)
        JotUtil.menuItemTrackingConfig.Track(My.Forms.Form1.HideRecentFilesListToolStripMenuItem)
        JotUtil.menuItemTrackingConfig.Track(My.Forms.Form1.AddProgramShortcutToExplorersContextmenuToolStripMenuItem)
        JotUtil.toolStripComboBoxTrackingConfig.Track(My.Forms.Form1.ToolStripComboBoxFontSize)
    End Sub

#End Region

End Module
