' This source-code is freely distributed as part of "DevCase Class Library .NET Developers".
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

Imports System.ComponentModel
Imports System.Drawing.Design
Imports System.IO

Imports DevCase.Core.Imaging.Tools
Imports DevCase.Core.IO
Imports DevCase.Interop.Unmanaged.Win32

#End Region

#Region " ShortcutFileInfoIconIndexEditor "

Namespace DevCase.Core.Design

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Provides a user interface for selecting the value of <see cref="ShortcutFileInfo.IconIndex"/> property.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <seealso cref="IconEditor"/>
    ''' ----------------------------------------------------------------------------------------------------
    Friend NotInheritable Class ShortcutFileInfoIconIndexEditor : Inherits IconEditor

#Region " Private Fields "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' A reference to <see cref="ShortcutFileInfo.Icon"/> property.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private iconPath As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' A reference to <see cref="ShortcutFileInfo.Target"/> property.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private target As String

#End Region

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Paints a representative value of the given object to the provided canvas.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="e">
        ''' What to paint and where to paint it.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Public Overrides Sub PaintValue(e As PaintValueEventArgs)
            Dim ico As Icon = Nothing
            Dim index As Integer = CInt(e.Value)

            Dim isIconPathEmpty As Boolean = String.IsNullOrWhiteSpace(Me.iconPath)
            Dim iconFileExist As Boolean = File.Exists(Me.iconPath)

            Dim isTargetEmpty As Boolean = String.IsNullOrWhiteSpace(Me.target)
            Dim targetFileExist As Boolean = File.Exists(Me.target)
            Dim targetDirExist As Boolean = System.IO.Directory.Exists(Me.target)

            ' Try extract icon from icon file path.
            If Not isIconPathEmpty AndAlso (iconFileExist) Then
                Try
                    ico = ImageUtil.ExtractIconFromExecutableFile(Me.iconPath, index)
                Catch
                End Try
            End If

            ' Try extract icon from target file path.
            If (ico Is Nothing) AndAlso Not (isTargetEmpty) AndAlso (targetFileExist) Then
                Try
                    ico = ImageUtil.ExtractIconFromExecutableFile(Me.target, index)
                Catch
                End Try
            End If

            ' Try extract icon from target file extension.
            If (ico Is Nothing) AndAlso Not (isTargetEmpty) AndAlso (targetFileExist) Then
                Try
                    Dim ext As String = Path.GetExtension(Me.target)
                    ico = ImageUtil.ExtractIconFromFileExtension(ext)
                Catch
                End Try
            End If

            ' Try extract icon from target directory path.
            If (ico Is Nothing) AndAlso Not (isTargetEmpty) AndAlso (targetDirExist) Then
                Try
                    ico = ImageUtil.ExtractIconFromDirectory(Me.target)
                Catch
                End Try
            End If

            ' Set default (null) file icon.
            If (ico Is Nothing) Then
                Try
                    ico = ImageUtil.ExtractIconFromExecutableFile("Shell32.dll", 0)
                Catch
                End Try
            End If

            ' Draw the icon.
            If (ico IsNot Nothing) Then
                e.Graphics.DrawIcon(ico, e.Bounds)
            End If

            MyBase.PaintValue(e)
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Edits the given object value using the editor style provided by the 
        ''' <see cref="IconEditor.GetEditStyle(ITypeDescriptorContext)"/> method.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="context">
        ''' A type descriptor context that can be used to provide additional context information.
        ''' </param>
        ''' 
        ''' <param name="provider">
        ''' A service provider object through which editing services may be obtained.
        ''' </param>
        ''' 
        ''' <param name="value">
        ''' An instance of the value being edited.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The new value of the object. 
        ''' <para></para>
        ''' If the value of the object has not changed, this should return the same object it was passed.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Overrides Function EditValue(context As ITypeDescriptorContext, provider As IServiceProvider, value As Object) As Object

            Me.UpdateFields(context)

            Dim path As String = Me.iconPath
            If String.IsNullOrEmpty(path) Then
                path = Me.target
            End If

            If File.Exists(path) Then
                If path.EndsWith(".ico", StringComparison.OrdinalIgnoreCase) OrElse
                    path.EndsWith(".icl", StringComparison.OrdinalIgnoreCase) OrElse
                    path.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) OrElse
                    path.EndsWith(".dll", StringComparison.OrdinalIgnoreCase) Then

                    Dim iconIndex As Integer
                    If NativeMethods.PickIconDlg(Process.GetCurrentProcess.MainWindowHandle(), path, 0, iconIndex) Then
                        value = iconIndex
                    End If
                End If
            End If

            Return value

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves the editing style of the <see cref="IconEditor.EditValue"/> method.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="context">
        ''' A type descriptor context that can be used to provide additional context information.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' One of the <see cref="UITypeEditorEditStyle"/> values indicating the provided editing style.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Overrides Function GetEditStyle(context As ITypeDescriptorContext) As UITypeEditorEditStyle
            Me.UpdateFields(context)
            Return UITypeEditorEditStyle.Modal
        End Function

#End Region

#Region " Private Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Updates the <see cref="ShortcutFileInfoIconIndexEditor.iconPath"/> and 
        ''' <see cref="ShortcutFileInfoIconIndexEditor.target"/> values.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="context">
        ''' A type descriptor context that can be used to provide additional context information.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub UpdateFields(context As ITypeDescriptorContext)
            Dim lnk As ShortcutFileInfo = DirectCast(context.Instance, ShortcutFileInfo)
            Me.iconPath = lnk.Icon
            Me.target = lnk.Target
        End Sub

#End Region

    End Class

End Namespace

#End Region
