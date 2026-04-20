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

Imports System.Runtime.CompilerServices

#End Region

#Region " MenuStrip Extensions "

' ReSharper disable once CheckNamespace

Namespace DevCase.Extensions.MenuStripExtensions

    ''' <summary>
    ''' Provides extension methods for <see cref="MenuStrip"/> type.
    ''' </summary>
    <HideModuleName>
    Public Module MenuStripExtensions

#Region " Public Extension Methods "

        ''' <summary>
        ''' Iterates through all the items within the source <see cref="MenuStrip"/> control, 
        ''' optionally recursively, and performs the specified action on each item.
        ''' </summary>
        '''
        ''' <param name="menuStrip">
        ''' The <see cref="MenuStrip"/> control whose items are to be iterated.
        ''' </param>
        ''' 
        ''' <param name="recursive">
        ''' <see langword="True"/> to iterate recursively through all items 
        ''' (i.e., iterate the child items of child items); otherwise, <see langword="False"/>.
        ''' </param>
        ''' 
        ''' <param name="action">
        ''' The action to perform on each item.
        ''' </param>
        <Extension>
        <DebuggerStepThrough>
        Public Sub ForEachItem(menuStrip As MenuStrip, recursive As Boolean, action As Action(Of ToolStripItem))

            ToolStripExtensions.ForEachItem(Of ToolStripItem)(menuStrip, recursive, action)

        End Sub

        ''' <summary>
        ''' Iterates through all the items of the specified type within the source <see cref="MenuStrip"/> control, 
        ''' optionally recursively, and performs the specified action on each item.
        ''' </summary>
        '''
        ''' <typeparam name="T">
        ''' The type of items to iterate through.
        ''' </typeparam>
        ''' 
        ''' <param name="menuStrip">
        ''' The <see cref="MenuStrip"/> control whose items are to be iterated.
        ''' </param>
        ''' 
        ''' <param name="recursive">
        ''' <see langword="True"/> to iterate recursively through all items 
        ''' (i.e., iterate the child items of child items); otherwise, <see langword="False"/>.
        ''' </param>
        ''' 
        ''' <param name="action">
        ''' The action to perform on each item.
        ''' </param>
        <Extension>
        <DebuggerStepThrough>
        Public Sub ForEachItem(Of T As ToolStripItem)(menuStrip As MenuStrip, recursive As Boolean, action As Action(Of T))

            ToolStripExtensions.ForEachItem(menuStrip, recursive, action)

        End Sub

#End Region

    End Module

End Namespace

#End Region
