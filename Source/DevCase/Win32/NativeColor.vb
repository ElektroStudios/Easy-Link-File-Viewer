' ***********************************************************************
' Author   : ElektroStudios
' Modified : 24-September-2024
' ***********************************************************************

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.Runtime.InteropServices

#End Region

#Region " NativeColor "

' ReSharper disable once CheckNamespace

Namespace DevCase.Win32.Structures

    ''' <summary>
    ''' Represents an RGB color (COLORREF) in the form: <c>0x00bbggrr</c>.
    ''' </summary>
    ''' 
    ''' <remarks>
    ''' <see href="https://learn.microsoft.com/en-us/windows/win32/gdi/colorref"/>
    ''' </remarks>
    <StructLayout(LayoutKind.Explicit, Size:=4)>
    Public Structure NativeColor : Implements IEquatable(Of NativeColor)

#Region " Fields "

        ''' <summary>
        ''' The DWORD value.
        ''' </summary>
        <FieldOffset(0)>
        Private ReadOnly Value As Integer

        ''' <summary>
        ''' The intensity of the red color.
        ''' </summary>
        <FieldOffset(0)>
        Public R As Byte

        ''' <summary>
        ''' The intensity of the green color.
        ''' </summary>
        <FieldOffset(1)>
        Public G As Byte

        ''' <summary>
        ''' The intensity of the blue color.
        ''' </summary>
        <FieldOffset(2)>
        Public B As Byte

        ''' <summary>
        ''' The transparency (alpha channel).
        ''' </summary>
        <FieldOffset(3)>
        Public A As Byte

        ''' <summary>
        ''' Represents the default color value. 
        ''' <para></para>
        ''' Used to specify that the default system color should be used.
        ''' </summary>
        Private Const CLR_DEFAULT As Integer = &HFF000000

        ''' <summary>
        ''' Represents a invalid color value. 
        ''' <para></para>
        ''' Used to specify that there is an error.
        ''' </summary>
        Private Const CLR_INVALID As Integer = &HFFFFFFFF

        ''' <summary>
        ''' Represents the static reference for a <see cref="NativeColor"/> with the default color (CLR_DEFAULT).
        ''' </summary>
        Public Shared ReadOnly [Default] As New NativeColor(NativeColor.CLR_DEFAULT)

        ''' <summary>
        ''' Represents the static reference for a <see cref="NativeColor"/> that is invalid (CLR_INVALID).
        ''' </summary>
        Public Shared ReadOnly Invalid As New NativeColor(NativeColor.CLR_INVALID)

#End Region

#Region " Constructors "

        ''' <summary>
        ''' Initializes a new instance of the <see cref="NativeColor"/> structure.
        ''' </summary>
        ''' 
        ''' <param name="r">
        ''' The intensity of the red color.
        ''' </param>
        ''' 
        ''' <param name="g">
        ''' The intensity of the green color.
        ''' </param>
        ''' 
        ''' <param name="b">
        ''' The intensity of the blue color.
        ''' </param>
        Public Sub New(r As Byte, g As Byte, b As Byte)
            Me.Value = 0
            Me.A = 0
            Me.R = r
            Me.G = g
            Me.B = b
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the <see cref="NativeColor"/> structure.
        ''' </summary>
        ''' 
        ''' <param name="value">
        ''' The packed DWORD value.
        ''' </param>
        Public Sub New(value As Integer)
            Me.R = 0
            Me.G = 0
            Me.B = 0
            Me.A = 0
            Me.Value = value And &HFFFFFF
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the <see cref="NativeColor"/> structure.
        ''' </summary>
        ''' 
        ''' <param name="color">
        ''' The color.
        ''' </param>
        Public Sub New(color As Color)
            Me.New(color.R, color.G, color.B)
            If color = Color.Transparent Then
                Me.Value = NativeColor.CLR_INVALID
            End If
        End Sub

#End Region

#Region " Equality Operators "

        ''' <summary>
        ''' Overloads the equality operator to compare two <see cref="NativeColor"/> values.
        ''' </summary>
        ''' 
        ''' <param name="left">
        ''' The first <see cref="NativeColor"/> to compare.
        ''' </param>
        ''' 
        ''' <param name="right">
        ''' The second <see cref="NativeColor"/> to compare.
        ''' </param>
        ''' 
        ''' <returns>
        ''' <see langword="True"/> if the colors are equal; 
        ''' otherwise, <see langword="False"/>.
        ''' </returns>
        Public Shared Operator =(left As NativeColor, right As NativeColor) As Boolean
            Return left.Equals(right)
        End Operator

        ''' <summary>
        ''' Overloads the inequality operator to compare two <see cref="NativeColor"/> values.
        ''' </summary>
        ''' 
        ''' <param name="left">
        ''' The first <see cref="NativeColor"/> to compare.
        ''' </param>
        ''' 
        ''' <param name="right">
        ''' The second <see cref="NativeColor"/> to compare.
        ''' </param>
        ''' 
        ''' <returns>
        ''' <see langword="True"/> if the colors are not equal; 
        ''' otherwise, <see langword="False"/>.
        ''' </returns>
        Public Shared Operator <>(left As NativeColor, right As NativeColor) As Boolean
            Return Not left = right
        End Operator

#End Region

#Region " Conversion Operators "

        ''' <summary>
        ''' Performs an implicit conversion from <see cref="NativeColor"/> to <see cref="Color"/>.
        ''' </summary>
        ''' 
        ''' <param name="color">
        ''' The color.
        ''' </param>
        ''' 
        ''' <returns>
        ''' The result of the conversion.
        ''' </returns>
        Public Shared Widening Operator CType(color As NativeColor) As Color
            Return If(color.Value = NativeColor.CLR_INVALID, Drawing.Color.Transparent, Drawing.Color.FromArgb(color.R, color.G, color.B))
        End Operator

        ''' <summary>
        ''' Performs an implicit conversion from <see cref="Color"/> to <see cref="NativeColor"/>.
        ''' </summary>
        ''' 
        ''' <param name="color">
        ''' The color.
        ''' </param>
        ''' 
        ''' <returns>
        ''' The result of the conversion.
        ''' </returns>
        Public Shared Widening Operator CType(color As Color) As NativeColor
            Return New NativeColor(color)
        End Operator

        ''' <summary>
        ''' Performs an implicit conversion from <see cref="ValueTuple(Of Byte, Byte, Byte)"/> 
        ''' to <see cref="NativeColor"/>.
        ''' </summary>
        ''' 
        ''' <param name="color">
        ''' The color.
        ''' </param>
        ''' 
        ''' <returns>
        ''' The result of the conversion.
        ''' </returns>
        Public Shared Widening Operator CType(color As (Byte, Byte, Byte)) As NativeColor
            Return New NativeColor(color.Item1, color.Item2, color.Item3)
        End Operator

        ''' <summary>
        ''' Performs an implicit conversion from <see cref="NativeColor"/> to <see cref="UInteger"/>.
        ''' </summary>
        ''' 
        ''' <param name="color">
        ''' The color.
        ''' </param>
        ''' 
        ''' <returns>
        ''' The result of the conversion.
        ''' </returns>
        Public Shared Widening Operator CType(color As NativeColor) As Integer
            Return color.Value
        End Operator

        ''' <summary>
        ''' Performs an implicit conversion from <see cref="UInteger"/> to <see cref="NativeColor"/>.
        ''' </summary>
        ''' 
        ''' <param name="value">
        ''' The value.
        ''' </param>
        ''' 
        ''' <returns>
        ''' The result of the conversion.
        ''' </returns>
        Public Shared Widening Operator CType(value As UInteger) As NativeColor
            Return New NativeColor(CInt(value))
        End Operator

        ''' <summary>
        ''' Performs an implicit conversion from <see cref="Integer"/> to <see cref="NativeColor"/>.
        ''' </summary>
        ''' 
        ''' <param name="value">
        ''' The value.
        ''' </param>
        ''' 
        ''' <returns>
        ''' The result of the conversion.
        ''' </returns>
        Public Shared Widening Operator CType(value As Integer) As NativeColor
            Return New NativeColor(value)
        End Operator

#End Region

#Region " Public Methods "

        ''' <summary>
        ''' Determines whether the specified <see cref="Object"/>, is equal to this instance.
        ''' </summary>
        ''' 
        ''' <param name="obj">
        ''' The <see cref="Object"/> to compare with this instance.
        ''' </param>
        ''' 
        ''' <returns>
        ''' <see langword="True"/> if the specified <see cref="Object"/> is equal to this instance; 
        ''' otherwise, <see langword="False"/>.
        ''' </returns>
        Public Overrides Function Equals(obj As Object) As Boolean
            Return TypeOf obj Is NativeColor AndAlso Equals(DirectCast(obj, NativeColor))
        End Function

        ''' <summary>
        ''' Determines whether the specified <see cref="NativeColor"/> is equal to the current instance.
        ''' </summary>
        ''' 
        ''' <param name="color">
        ''' The <see cref="NativeColor"/> to compare with the current instance.
        ''' </param>
        ''' 
        ''' <returns>
        ''' <see langword="True"/> if the specified <see cref="NativeColor"/> is equal to the current instance; 
        ''' otherwise, <see langword="False"/>.
        ''' </returns>
        Public Overloads Function Equals(color As NativeColor) As Boolean Implements IEquatable(Of NativeColor).Equals
            Return color.A = Me.A AndAlso color.B = Me.B AndAlso color.G = Me.G AndAlso color.R = Me.R
        End Function

        ''' <summary>
        ''' Returns a hash code for this instance.
        ''' </summary>
        ''' 
        ''' <returns>
        ''' A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
        ''' </returns>
        Public Overrides Function GetHashCode() As Integer
            Return Me.ToArgb()
        End Function

        ''' <summary>
        ''' Converts the current <see cref="NativeColor"/> to a 32-bit ARGB value.
        ''' </summary>
        ''' 
        ''' <returns>
        ''' A 32-bit integer representing the ARGB value of the current <see cref="NativeColor"/>.
        ''' </returns>
        Public Function ToArgb() As Integer
            Return Me.Value
        End Function

        ''' <summary>
        ''' Converts the current <see cref="NativeColor"/> 
        ''' to a string that represents the corresponding <see cref="Color"/> object.
        ''' </summary>
        ''' 
        ''' <returns>
        ''' A string representation of the current <see cref="NativeColor"/> as a <see cref="Color"/>.
        ''' </returns>
        Public Overrides Function ToString() As String
            Return CType(Me, Color).ToString()
        End Function

#End Region

    End Structure

End Namespace

#End Region
