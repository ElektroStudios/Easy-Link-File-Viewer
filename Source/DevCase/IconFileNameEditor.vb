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

Imports System.Windows.Forms.Design

#End Region

#Region " IconFileNameEditor "

Namespace DevCase.Core.Design

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Provides a user interface for selecting a icon file name.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <seealso cref="FileNameEditor"/>
    ''' ----------------------------------------------------------------------------------------------------
    Friend Class IconFileNameEditor : Inherits FileNameEditor

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Initializes a new instance of the <see cref="IconFileNameEditor"/> class.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Sub New()
            MyBase.New()
        End Sub

#End Region

#Region " Private Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Initializes the open file dialog when it is created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="dlg">
        ''' The <see cref="OpenFileDialog"/> to use to select a file name.
        ''' </param>
        Protected Overrides Sub InitializeDialog(dlg As OpenFileDialog)
            MyBase.InitializeDialog(dlg)

            With dlg
                .Multiselect = False
                .RestoreDirectory = True
                .DereferenceLinks = True
                .Filter = "Icon Files (*.ico;*.icl;*.exe;*.dll)|*.ico;*.icl;*.exe;*.dll|Icons|*.ico|Libraries|*.dll|Programs|*.exe"
                .FilterIndex = 1
                .SupportMultiDottedExtensions = True
            End With
        End Sub

#End Region

    End Class

End Namespace

#End Region
