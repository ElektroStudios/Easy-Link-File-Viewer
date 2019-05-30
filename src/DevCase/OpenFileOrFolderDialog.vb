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

Option Explicit On
Option Strict On
Option Infer Off

#End Region

#Region " Imports "

Imports System.ComponentModel
Imports System.IO

#End Region

#Region " OpenFileOrFolderDialog "

Namespace DevCase.UserControls.Controls

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Displays a dialog box that prompts the user to open a (single) file or folder.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' Original source-code: 
    ''' <para></para>
    ''' <see href="https://www.codeproject.com/Articles/44914/Select-file-or-folder-from-the-same-dialog"/>
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <seealso cref="CommonDialog" />
    ''' ----------------------------------------------------------------------------------------------------
    <DisplayName(NameOf(OpenFileOrFolderDialog))>
    <Description("Displays a standard dialog box that prompts the user to open a (single) file or folder.")>
    <DesignTimeVisible(True)>
    <DesignerCategory(NameOf(Component))>
    <ToolboxBitmap(GetType(Component), "Component.bmp")>
    <ToolboxItemFilter("System.Windows.Forms")>
    <DefaultEvent(NameOf(OpenFileOrFolderDialog.ItemOk))>
    <DefaultProperty(NameOf(OpenFileOrFolderDialog.ItemName))>
    Friend Class OpenFileOrFolderDialog : Inherits CommonDialog

#Region " Private Fields "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The underlying <see cref="OpenFileDialog"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Protected WithEvents Dialog As OpenFileDialog

#End Region

#Region " Properties "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the custom places collection for this <see cref="OpenFileOrFolderDialog"/> instance.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The custom places collection for this <see cref="OpenFileOrFolderDialog"/> instance.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Browsable(False)>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
        Public ReadOnly Property CustomPlaces As FileDialogCustomPlacesCollection
            <DebuggerStepThrough>
            Get
                Return Me.Dialog.CustomPlaces
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the dialog box title.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The dialog box title. The default value is an empty string ("").
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <DefaultValue("")>
        <Localizable(True)>
        <Description("The dialog box title.")>
        Public Property Title As String
            <DebuggerStepThrough>
            Get
                Return Me.Dialog.Title
            End Get
            <DebuggerStepThrough>
            Set(value As String)
                Me.Dialog.Title = value
            End Set
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value indicating whether the Help button is displayed in the dialog box.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if the dialog box includes a help button; 
        ''' otherwise, <see langword="False"/>. 
        ''' <para></para>
        ''' The default value is <see langword="False"/>.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <DefaultValue(False)>
        <Description("A value indicating whether the Help button is displayed in the dialog box.")>
        Public Property ShowHelp As Boolean
            <DebuggerStepThrough>
            Get
                Return Me.Dialog.ShowHelp
            End Get
            <DebuggerStepThrough>
            Set(value As Boolean)
                Me.Dialog.ShowHelp = value
            End Set
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value indicating whether the dialog box restores the directory 
        ''' to the previously selected directory before closing.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if the dialog box restores the current directory 
        ''' to the previously selected directory if the user changed the directory while searching for files; 
        ''' otherwise, <see langword="False"/>. 
        ''' <para></para>
        ''' The default value is <see langword="False"/>.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <DefaultValue(False)>
        <Description("A value indicating whether the dialog box restores the directory to the previously selected directory before closing.")>
        Public Property RestoreDirectory As Boolean
            <DebuggerStepThrough>
            Get
                Return Me.Dialog.RestoreDirectory
            End Get
            <DebuggerStepThrough>
            Set(value As Boolean)
                Me.Dialog.RestoreDirectory = value
            End Set
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the initial directory displayed by the dialog box.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The initial directory displayed by the dialog box. The default is an empty string ("").
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <DefaultValue("")>
        <Description("The initial directory displayed by the dialog box.")>
        Public Property InitialDirectory As String
            <DebuggerStepThrough>
            Get
                Return Me.Dialog.InitialDirectory
            End Get
            <DebuggerStepThrough>
            Set(value As String)
                Me.Dialog.InitialDirectory = value
            End Set
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value indicating whether this dialog box should 
        ''' automatically upgrade appearance and behavior when running on Windows Vista.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if this System.Windows.Forms.FileDialog instance 
        ''' should automatically upgrade appearance and behavior when running on Windows Vista; 
        ''' otherwise, <see langword="False"/>. 
        ''' <para></para>
        ''' The default value is <see langword="True"/>.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <DefaultValue(True)>
        <Description("A value indicating whether this dialog box should automatically upgrade appearance and behavior when running on Windows Vista.")>
        Public Property AutoUpgradeEnabled As Boolean
            <DebuggerStepThrough>
            Get
                Return Me.Dialog.AutoUpgradeEnabled
            End Get
            <DebuggerStepThrough>
            Set(value As Boolean)
                Me.Dialog.AutoUpgradeEnabled = value
            End Set
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value indicating whether the dialog box returns the location of the 
        ''' file referenced by the shortcut or whether it returns the location of the shortcut (.lnk).
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if the dialog box returns the location of the file referenced by the shortcut;
        ''' otherwise, <see langword="False"/>. 
        ''' <para></para>
        ''' The default value is <see langword="False"/>.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <DefaultValue(True)>
        <Description("A value indicating whether the dialog box returns the location of the file referenced by the shortcut or whether it returns the location of the shortcut (.lnk).")>
        Public Property DereferenceLinks As Boolean
            <DebuggerStepThrough>
            Get
                Return Me.Dialog.DereferenceLinks
            End Get
            <DebuggerStepThrough>
            Set(value As Boolean)
                Me.Dialog.DereferenceLinks = value
            End Set
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a string containing the name of the item selected in the dialog box.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>The name of the item selected in the dialog box.
        ''' <para></para>
        ''' The default value is an empty string ("").
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <DefaultValue("")>
        <Description("A string containing the name of the item selected in the dialog box.")>
        Public Property ItemName As String
            <DebuggerStepThrough>
            Get
                Try
                    If Not String.IsNullOrWhiteSpace(Me.Dialog.FileName) AndAlso
                        (Me.Dialog.FileName.EndsWith("Folder Selection.", StringComparison.OrdinalIgnoreCase) OrElse
                        Not File.Exists(Me.Dialog.FileName)) AndAlso
                        Not Directory.Exists(Me.Dialog.FileName) Then
                        Return Path.GetDirectoryName(Me.Dialog.FileName)
                    Else
                        Return Me.Dialog.FileName

                    End If

                Catch
                    Return Me.Dialog.FileName
                End Try
            End Get
            <DebuggerStepThrough>
            Set(ByVal value As String)
                Me.Dialog.FileName = value
            End Set
        End Property

#End Region

#Region " Events "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Occurs when the user clicks on the Open button on the dialog box to select a file or folder.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Description("Occurs when the user clicks on the Open button on the dialog box to select a file or folder.")>
        Public Event ItemOk As CancelEventHandler

#End Region

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Initializes a new instance of the <see cref="OpenFileOrFolderDialog"/> class.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Sub New()
            ' Set validate names to false. Otherwise, windows will not let you select "Folder Selection."
            Me.Dialog = New OpenFileDialog With {
                .Multiselect = False,
                .ValidateNames = False,
                .CheckFileExists = False,
                .CheckPathExists = True
            }
        End Sub

#End Region

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Runs a common dialog box with a default owner.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see cref="DialogResult.OK"/> if the user clicks OK in the dialog box; 
        ''' otherwise, <see cref="DialogResult.Cancel"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shadows Function ShowDialog() As DialogResult
            Return Me.ShowDialog(Nothing)
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Runs a common dialog box with a default owner.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="owner">
        ''' Any object that implements <see cref="IWin32Window"/> that represents the 
        ''' top-level window that will own the modal dialog box.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see cref="DialogResult.OK"/> if the user clicks OK in the dialog box; 
        ''' otherwise, <see cref="DialogResult.Cancel"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shadows Function ShowDialog(ByVal owner As IWin32Window) As DialogResult
            ' Set initial directory (used when 'Me.dialog.FileName' is set from outside)
            If Not (String.IsNullOrWhiteSpace(Me.Dialog.FileName)) Then
                Try
                    If Directory.Exists(Me.Dialog.FileName) Then
                        Me.Dialog.InitialDirectory = Me.Dialog.FileName
                    Else
                        Me.Dialog.InitialDirectory = Path.GetDirectoryName(Me.Dialog.FileName)
                    End If

                Catch
                    ' Do nothing
                End Try
            End If

            ' Always default to "Folder Selection."
            Me.Dialog.FileName = "Folder Selection."

            If (owner Is Nothing) Then
                Return Me.Dialog.ShowDialog()
            Else
                Return Me.Dialog.ShowDialog(owner)
            End If
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Resets the properties of a common dialog box to their default values.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Overrides Sub Reset()
            Me.Dialog.Reset()
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Returns a string version of this object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A string version of this object.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Overrides Function ToString() As String
            Return Me.Dialog.ToString()
        End Function

#End Region

#Region " Event Invocators "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Raises the System.Windows.Forms.FileDialog.FileOk event.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="e">
        ''' A <see cref="CancelEventArgs"/> that contains the event data.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Protected Sub OnItemOk(ByVal e As CancelEventArgs)
            If (Me.ItemOkEvent IsNot Nothing) Then
                RaiseEvent ItemOk(Me, e)
            End If
        End Sub

#End Region

#Region " Event Handlers "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Handles the <see cref="OpenFileDialog.FileOk"/> event of the <see cref="OpenFileOrFolderDialog.Dialog"/> component.
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
        <DebuggerStepThrough>
        Private Sub Dialog_OnFileOk(ByVal sender As Object, ByVal e As CancelEventArgs) Handles Dialog.FileOk
            Me.OnItemOk(e)
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Handles the <see cref="OpenFileDialog.HelpRequest"/> event of the <see cref="OpenFileOrFolderDialog.Dialog"/> component.
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
        <DebuggerStepThrough>
        Private Sub Dialog_OnFileOk(ByVal sender As Object, ByVal e As EventArgs) Handles Dialog.HelpRequest
            Me.OnHelpRequest(e)
        End Sub

#End Region

#Region " Private Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Specifies a common dialog box.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hwndOwner">
        ''' A value that represents the window handle of the owner window for the common dialog box.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see langword="True"/> if the dialog box was successfully run; otherwise, <see langword="False"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepperBoundary>
        Protected Overrides Function RunDialog(ByVal hwndOwner As IntPtr) As Boolean
            Return True
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Releases the unmanaged resources used by the <see cref="T:System.ComponentModel.Component" /> and optionally releases the managed resources.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="disposing">
        ''' <see langword="True"/> to release both managed and unmanaged resources; 
        ''' <see langword="false"/> to release only unmanaged resources.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepperBoundary>
        Protected Overrides Sub Dispose(disposing As Boolean)
            Me.Dialog?.Dispose()
            MyBase.Dispose(disposing)
        End Sub

#End Region

    End Class

End Namespace

#End Region