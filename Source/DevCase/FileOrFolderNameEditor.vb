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
Imports DevCase.UserControls.Controls

#End Region

#Region " FileOrFolderNameEditor "

Namespace DevCase.Core.Design

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Provides a user interface for selecting a file or folder name.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <seealso cref="UITypeEditor"/>
    ''' ----------------------------------------------------------------------------------------------------
    Friend Class FileOrFolderNameEditor : Inherits UITypeEditor

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Initializes a new instance of the <see cref="FileOrFolderNameEditor"/> class.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Sub New()
            MyBase.New()
        End Sub

#End Region

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the editor style used by the <see cref="UITypeEditor.EditValue(IServiceProvider, Object)"/> method.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="context">
        ''' An <see cref="ITypeDescriptorContext"/> that can be used to gain additional context information.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A <see cref="UITypeEditorEditStyle"/> value that indicates the style of editor used 
        ''' by the <see cref="UITypeEditor.EditValue(IServiceProvider, Object)"/> method. 
        ''' <para></para>
        ''' If the <see cref="UITypeEditor"/> does not support this method, 
        ''' then <see cref="UITypeEditor.GetEditStyle"/> will return <see cref="UITypeEditorEditStyle.None"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Overrides Function GetEditStyle(context As ITypeDescriptorContext) As UITypeEditorEditStyle
            Return UITypeEditorEditStyle.Modal
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Edits the specified object's value using the editor style indicated by the <see cref="UITypeEditor.GetEditStyle"/> method.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="context">
        ''' An <see cref="ITypeDescriptorContext"/> that can be used to gain additional context information.
        ''' </param>
        ''' 
        ''' <param name="provider">
        ''' An <see cref="IServiceProvider"/> that this editor can use to obtain services.
        ''' </param>
        ''' 
        ''' <param name="value">
        ''' The object to edit.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The new value of the object. 
        ''' <para></para>
        ''' If the value of the object has not changed, this should return the same object it was passed.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Overrides Function EditValue(context As ITypeDescriptorContext, provider As IServiceProvider, value As Object) As Object

            Using dlg As New OpenFileOrFolderDialog() With {
                .Title = "Select a file or folder...",
                .AutoUpgradeEnabled = True,
                .DereferenceLinks = True,
                .RestoreDirectory = True,
                .ShowHelp = False
            }

                If (dlg.ShowDialog = DialogResult.OK) Then
                    Return dlg.ItemName
                End If
            End Using

            Return MyBase.EditValue(context, provider, value)

        End Function

#End Region

    End Class

End Namespace

#End Region
