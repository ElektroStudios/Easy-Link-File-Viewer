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

#Region " ContainerControl Extensions "

' ReSharper disable once CheckNamespace

Namespace DevCase.Extensions.ContainerControlExtensions

    ''' <summary>
    ''' Provides extension methods for <see cref="ContainerControl"/>.
    ''' </summary>
    <HideModuleName>
    Public Module ContainerControlExtensions

#Region " Public Extension Methods "

        ''' <summary>
        ''' Iterates through all controls of the specified type within a parent <see cref="ContainerControl"/>, 
        ''' optionally recursively, and performs the specified action on each control.
        ''' </summary>
        '''
        ''' <typeparam name="T">
        ''' The type of child controls to iterate through.
        ''' </typeparam>
        ''' 
        ''' <param name="container">
        ''' The parent <see cref="ContainerControl"/> whose child controls are to be iterated.
        ''' </param>
        ''' 
        ''' <param name="recursive">
        ''' <see langword="True"/> to iterate recursively through all child controls 
        ''' (i.e., iterate the child controls of child controls); otherwise, <see langword="False"/>.
        ''' </param>
        ''' 
        ''' <param name="action">
        ''' The action to perform on each control.
        ''' </param>
        <Extension>
        <DebuggerStepThrough>
        Public Sub ForEachControl(Of T As Control)(container As ContainerControl, recursive As Boolean, action As Action(Of T))

            ControlExtensions.ForEachControl(container, recursive, action)

        End Sub

#End Region

    End Module

End Namespace

#End Region
