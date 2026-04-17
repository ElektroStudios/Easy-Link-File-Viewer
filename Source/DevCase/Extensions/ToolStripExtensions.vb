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

#Region " ToolStrip Extensions "

' ReSharper disable once CheckNamespace

Namespace DevCase.Extensions.ToolStripExtensions

    ''' <summary>
    ''' Provides custom extension methods to use with <see cref="ToolStrip"/> class.
    ''' </summary>
    <HideModuleName>
    Public Module ToolStripExtensions

#Region " Public Extension Methods "

        ''' <summary>
        ''' Iterates through all the items of the specified type within the source <see cref="ToolStrip"/> control, 
        ''' optionally recursively, and performs the specified action on each item.
        ''' </summary>
        '''
        ''' <param name="toolStrip">
        ''' The <see cref="ToolStrip"/> control whose items are to be iterated.
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
        Public Sub ForEachItem(toolStrip As ToolStrip, recursive As Boolean, action As Action(Of ToolStripItem))

            ForEachItem(Of ToolStripItem)(toolStrip, recursive, action)

        End Sub

        ''' <summary>
        ''' Iterates through all the items of the specified type within the source <see cref="ToolStrip"/> control, 
        ''' optionally recursively, and performs the specified action on each item.
        ''' </summary>
        '''
        ''' <typeparam name="T">
        ''' The type of items to iterate through.
        ''' </typeparam>
        ''' 
        ''' <param name="toolStrip">
        ''' The <see cref="ToolStrip"/> control whose items are to be iterated.
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
        Public Sub ForEachItem(Of T As ToolStripItem)(toolStrip As ToolStrip, recursive As Boolean, action As Action(Of T))
            If action Is Nothing Then
                Throw New ArgumentNullException(paramName:=NameOf(action), "Action cannot be null.")
            End If

            Dim queue As New Queue(Of ToolStripItem)

            ' First level items iteration.
            For Each item As ToolStripItem In toolStrip.Items
                If recursive Then
                    queue.Enqueue(item)
                Else
                    If TypeOf item Is T Then
                        action.Invoke(DirectCast(item, T))
                    End If
                End If
            Next item

            ' Recursive items iteration.
            While queue.Any()
                Dim currentItem As ToolStripItem = queue.Dequeue()
                If TypeOf currentItem Is T Then
                    action.Invoke(DirectCast(currentItem, T))
                End If

                If TypeOf currentItem Is ToolStripDropDownItem Then
                    Dim dropDownItem As ToolStripDropDownItem = DirectCast(currentItem, ToolStripDropDownItem)
                    For Each subItem As ToolStripItem In dropDownItem.DropDownItems
                        queue.Enqueue(subItem)
                    Next subItem
                End If
            End While
        End Sub

#End Region

    End Module

End Namespace

#End Region
