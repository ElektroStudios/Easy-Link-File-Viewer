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

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.ComponentModel
Imports System.Reflection
Imports System.Runtime.CompilerServices

#End Region

#Region " PropertyGrid Extensions "

Namespace DevCase.Core.Extensions

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains custom extension methods to use with the <see cref="PropertyGrid"/> type.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    <HideModuleName>
    Friend Module PropertyGridExtensions

#Region " Public Extension Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Moves the splitter control of the source the <see cref="PropertyGrid"/> 
        ''' to adjust the columns width.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="propGrid">
        ''' The source <see cref="PropertyGrid"/>.
        ''' </param>
        ''' 
        ''' <param name="width">
        ''' The desired width.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        <Extension>
        <EditorBrowsable(EditorBrowsableState.Always)>
        Public Sub MoveSplitterTo(ByVal propGrid As PropertyGrid, ByVal width As Integer)

            If Not (propGrid.Visible) Then
                Throw New InvalidOperationException("The PropertyGrid must be visible in order to move its splitter control.")
            End If

            Dim fi As FieldInfo = propGrid.GetType().GetField("gridView", BindingFlags.Instance Or BindingFlags.NonPublic)
            Dim view As Control = TryCast(fi.GetValue(propGrid), Control) ' Internal Class Name: "PropertyGridView"
            Dim mi As MethodInfo = view.GetType().GetMethod("MoveSplitterTo", BindingFlags.Instance Or BindingFlags.NonPublic)

            mi.Invoke(view, {width})

        End Sub

#End Region

    End Module

End Namespace

#End Region
