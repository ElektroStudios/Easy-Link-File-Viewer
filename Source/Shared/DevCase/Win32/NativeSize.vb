' ***********************************************************************
' Author   : ElektroStudios
' Modified : 12-December-2015
' ***********************************************************************

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.Runtime.InteropServices

#End Region

#Region " NativeSize (Size) "

' ReSharper disable once CheckNamespace

Namespace DevCase.Win32.Structures

    ''' <summary>
    ''' Defines the width and height of a rectangle.
    ''' </summary>
    '''
    ''' <remarks>
    ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/dd145106%28v=vs.85%29.aspx"/>
    ''' </remarks>
    <DebuggerStepThrough>
    <StructLayout(LayoutKind.Sequential)>
    Public Structure NativeSize

#Region " Fields "

        ''' <summary>
        ''' Specifies the rectangle's width.
        ''' <para></para>
        ''' The units depend on which function uses this.
        ''' </summary>
        Public Width As Integer

        ''' <summary>
        ''' Specifies the rectangle's height.
        ''' <para></para>
        ''' The units depend on which function uses this.
        ''' </summary>
        Public Height As Integer

#End Region

#Region " Constructors "

        ''' <summary>
        ''' Initializes a new instance of the <see cref="NativeSize"/> struct.
        ''' </summary>
        '''
        ''' <param name="width">
        ''' The width.
        ''' </param>
        ''' 
        ''' <param name="height">
        ''' The height.
        ''' </param>
        Public Sub New(width As Integer, height As Integer)

            Me.Width = width
            Me.Height = height

        End Sub

        ''' <summary>
        ''' Initializes a new instance of the <see cref="NativeSize"/> struct.
        ''' </summary>
        '''
        ''' <param name="size">
        ''' A <see cref="Size"/> structure that contains the width and height.
        ''' </param>
        Public Sub New(size As Size)

            Me.New(size.Width, size.Height)

        End Sub

#End Region

#Region " Operator Conversions "

        ''' <summary>
        ''' Performs an implicit conversion from <see cref="NativeSize"/> to <see cref="Size"/>.
        ''' </summary>
        '''
        ''' <param name="size">
        ''' The <see cref="NativeSize"/>.
        ''' </param>
        '''
        ''' <returns>
        ''' The resulting <see cref="Size"/>.
        ''' </returns>
        Public Shared Widening Operator CType(size As NativeSize) As Size

            Return New Size(size.Width, size.Height)

        End Operator

        ''' <summary>
        ''' Performs an implicit conversion from <see cref="Size"/> to <see cref="NativeSize"/>.
        ''' </summary>
        '''
        ''' <param name="size">
        ''' The <see cref="Size"/>.
        ''' </param>
        '''
        ''' <returns>
        ''' The resulting <see cref="NativeSize"/>.
        ''' </returns>
        Public Shared Widening Operator CType(size As Size) As NativeSize

            Return New NativeSize(size.Width, size.Height)

        End Operator

#End Region

    End Structure

End Namespace

#End Region
