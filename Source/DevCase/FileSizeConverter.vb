' This source-code is freely distributed as part of "DevCase Class Library .NET Developers".
' 
' Maybe you would like to consider to buy this powerful set of libraries to support me. 
' You can do loads of things with my apis for a big amount of diverse thematics.
' 
' Here is a link to the purchase page:
' https://codecanyon.net/item/elektrokit-class-library-for-net/19260282
' 
' Thank you.

#Region " Usage Examples "

'<TypeConverter(GetType(FileSizeConverter))>
'<Browsable(True)>
'Public ReadOnly Property FileSize As Long = 2048 ' Bytes

'<TypeConverter(GetType(FileSizeConverter))>
'<Browsable(True)>
'Public ReadOnly Property FileSize As New Filesize(2048, SizeUnits.Byte)

#End Region

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.ComponentModel
Imports System.ComponentModel.Design.Serialization
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Text
Imports DevCase.Interop.Unmanaged.Win32

#End Region

#Region " FileSizeConverter "

Namespace DevCase.Core.Design

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Provides a type converter to converts a numeric value 
    ''' into a string that represents the number expressed as a size value in bytes, 
    ''' kilobytes, megabytes, gigabytes, petabytes or exabytes, depending on the size.
    ''' <para></para>
    ''' Conversion examples:
    ''' <para></para>
    ''' Input value -> Result string
    ''' <para></para>
    ''' 793   -> 793 Bytes
    ''' <para></para>
    ''' 1533  -> 1,49 KB
    ''' <para></para>
    ''' 2049 -> 2,00 KB
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Public Class Form1
    ''' 
    '''     Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    '''         Me.PropertyGrid1.SelectedObject = New TestClass
    '''     End Sub
    '''     
    ''' End Class
    ''' 
    ''' Public Class TestClass
    ''' 
    '''     &lt;TypeConverter(GetType(FileSizeConverter))&gt;
    '''     Public ReadOnly Property FileSize1 As Long = 1024 ' Bytes
    ''' 
    '''     &lt;TypeConverter(GetType(FileSizeConverter))&gt;
    '''     Public ReadOnly Property FileSize2 As New Filesize(1024, SizeUnits.Byte)
    '''     
    ''' End Class
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    Friend Class FileSizeConverter : Inherits TypeConverter

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Initializes a new instance of the <see cref="FileSizeConverter"/> class.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Sub New()
            MyBase.New()
        End Sub

#End Region

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Determines if this converter can convert an object in the given source type to the native type of the converter.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="context">
        ''' An <see cref="ITypeDescriptorContext" /> that provides a format context.
        ''' </param>
        ''' 
        ''' <param name="sourceType">
        ''' A <see cref="Type" /> that represents the type from which you want to convert.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see langword="true" /> if this converter can perform the operation; otherwise, <see langword="false" />.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Overrides Function CanConvertFrom(context As ITypeDescriptorContext, sourceType As Type) As Boolean

            Return False

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Returns whether this converter can convert the object to the specified type, using the specified context.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="context">
        ''' An <see cref="ITypeDescriptorContext"/> that provides a format context.
        ''' </param>
        ''' 
        ''' <param name="destinationType">
        ''' A <see cref="Type"/> that represents the type you want to convert to.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see langword="True"/> if this converter can perform the conversion; otherwise, <see langword="False"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Overloads Overrides Function CanConvertTo(context As ITypeDescriptorContext, destinationType As Type) As Boolean

            If (destinationType = GetType(InstanceDescriptor)) Then
                Return True
            End If

            Return MyBase.CanConvertTo(context, destinationType)

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Converts the specified object to another type.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="context">
        ''' An <see cref="ITypeDescriptorContext"/> that provides a format context.
        ''' </param>
        ''' 
        ''' <param name="culture">
        ''' A <see cref="CultureInfo"/> that specifies the culture to represent the number.
        ''' </param>
        ''' 
        ''' <param name="value">
        ''' The object to convert.
        ''' </param>
        ''' 
        ''' <param name="destinationType">
        ''' The type to convert the object to.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' An <see cref="Object"/> that represents the converted value.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Overrides Function ConvertTo(context As ITypeDescriptorContext, culture As CultureInfo, value As Object, destinationType As Type) As Object

            If (destinationType Is Nothing) Then
                Throw New ArgumentNullException("destinationType")
            End If

            If (destinationType Is GetType(String)) Then
                'If (TypeOf value Is DevCase.Core.IO.Filesize) Then  ' Return 'Filesize' string format.
                '    Return DirectCast(value, DevCase.Core.IO.Filesize).ToString()
                'Else
                Dim sb As New StringBuilder(16)
                Dim longValue As ULong
                If ULong.TryParse(value.ToString(), longValue) Then ' Return native string format.
                    Dim result As IntPtr = NativeMethods.StrFormatByteSize64A(longValue, sb, CUInt(sb.Capacity))
                    Dim win32Err As Integer = Marshal.GetLastWin32Error()

                    If (result <> IntPtr.Zero) Then
                        Return sb.ToString()
                    Else
                        Return New Win32Exception(win32Err).Message
                    End If
                Else ' Return original value.
                    Return value
                End If

                'End If
            End If

            Return MyBase.ConvertTo(context, culture, value, destinationType)

        End Function

#End Region

    End Class

End Namespace

#End Region
