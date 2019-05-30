' This source-code is freely distributed as part of "DevCase for .NET Framework".
' 
' Maybe you would like to consider to buy this powerful set of libraries to support me. 
' You can do loads of things with my apis for a big amount of diverse thematics.
' 
' Here is a link to the purchase page:
' https://codecanyon.net/item/elektrokit-class-library-for-net/19260282
' 
' Thank you.

#Region " Public Members Summary "

#Region " Constructors "

' New(Double, SizeUnits)

#End Region

#Region " Properties "

' Size(SizeUnits) As Double
' Size(SizeUnits, Integer, Opt: NumberFormatInfo) As String
' SizeRounded As Double
' SizeRounded(Integer, Opt: NumberFormatInfo) As String
' SizeUnit As SizeUnits
' SizeUnitNameShort As String
' SizeUnitNameLong As String

#End Region

#Region " Functions "

' ToString() As String

#End Region

#End Region

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.ComponentModel
Imports System.Globalization
Imports System.Xml.Serialization

Imports DevCase.Core.IO
Imports DevCase.Core.Design

#End Region

#Region " Filesize "

Namespace DevCase.Core.IO

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Defines a filesize.
    ''' <para></para>
    ''' Provides methods to round or convert a filesize between different units of size.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example that performs a simple conversion between units of size.
    ''' <code>
    ''' Dim fs As New Filesize(1073741824, Filesize.SizeUnits.Byte)
    ''' 
    ''' Dim b As Double = fs.Size(Filesize.SizeUnits.Byte)
    ''' Dim kb As Double = fs.Size(Filesize.SizeUnits.KiloByte)
    ''' Dim mb As Double = fs.Size(Filesize.SizeUnits.MegaByte)
    ''' Dim gb As Double = fs.Size(Filesize.SizeUnits.GigaByte)
    ''' Dim tb As Double = fs.Size(Filesize.SizeUnits.TeraByte)
    ''' Dim pb As Double = fs.Size(Filesize.SizeUnits.PetaByte)
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example that rounds a filesize in bytes, to its most approximated unit of size.
    ''' <code>
    ''' For Each sizeUnit As Filesize.SizeUnits In [Enum].GetValues(GetType(Filesize.SizeUnits))
    ''' 
    '''     Dim fsize As New Filesize(sizeUnit, Filesize.SizeUnits.Byte)
    ''' 
    '''     Dim stringFormat As String =
    '''         String.Format("{0} Bytes rounded to {1} {2}.",
    '''                       fsize.Size(Filesize.SizeUnits.Byte, CultureInfo.CurrentCulture.NumberFormat),
    '''                       fsize.SizeRounded(decimalPrecision:=2, numberFormatInfo:=Nothing),
    '''                       fsize.SizeUnitNameLong)
    ''' 
    '''     Console.WriteLine(stringFormat)
    ''' 
    ''' Next sizeUnit
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example that converts a Terabyte (1099511627776 Bytes) to other units of size.
    ''' <code>
    ''' Dim fsize As New Filesize(Filesize.SizeUnits.TeraByte, Filesize.SizeUnits.Byte)
    ''' 
    ''' For Each sizeUnit As Filesize.SizeUnits In [Enum].GetValues(GetType(Filesize.SizeUnits))
    ''' 
    '''     Dim stringFormat As String =
    '''         String.Format("{0} Bytes equals to {1} {2}.",
    '''                       fsize.Size(Filesize.SizeUnits.Byte, Nothing, CultureInfo.CurrentCulture.NumberFormat),
    '''                       fsize.Size(sizeUnit, 2, CultureInfo.CurrentCulture.NumberFormat),
    '''                       sizeUnit.ToString())
    ''' 
    '''     Console.WriteLine(stringFormat)
    ''' 
    ''' Next sizeUnit
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    <Serializable>
    <XmlRoot(NameOf(Filesize))>
    Friend NotInheritable Class Filesize

#Region " Private Fields "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The filesize, in Bytes.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private ReadOnly bytesB As Double

#End Region

#Region " Properties "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the filesize, in the specified unit of size.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="sizeUnit">
        ''' The unit of size.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The filesize.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public ReadOnly Property Size(ByVal sizeUnit As SizeUnits) As Double
            <DebuggerStepThrough>
            Get
                Return Me.Convert(size:=Me.bytesB, fromUnit:=SizeUnits.Byte, toUnit:=sizeUnit)
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the filesize, in the specified unit of size, using the specified numeric format.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="sizeUnit">
        ''' The unit of size.
        ''' </param>
        ''' 
        ''' <param name="numberFormatInfo">
        ''' A custom <see cref="NumberFormatInfo"/> format provider.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The filesize.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public ReadOnly Property Size(ByVal sizeUnit As SizeUnits,
                                      ByVal decimalPrecision As Integer,
                                      Optional ByVal numberFormatInfo As NumberFormatInfo = Nothing) As String
            <DebuggerStepThrough>
            Get
                If (numberFormatInfo Is Nothing) Then
                    numberFormatInfo = CultureInfo.CurrentCulture.NumberFormat
                End If
                Return Me.Size(sizeUnit).ToString(String.Format("N{0}", decimalPrecision), numberFormatInfo)
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the filesize, rounded using the most approximated unit of size.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The rounded filesize.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public ReadOnly Property SizeRounded As Double

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the filesize, rounded using the most approximated unit of size, with the specified decimal precision.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="decimalPrecision">
        ''' The decimal precision.
        ''' </param>
        ''' 
        ''' <param name="numberFormatInfo">
        ''' A <see cref="NumberFormatInfo"/> format provider.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The rounded value, with the specified decimal precision.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public ReadOnly Property SizeRounded(ByVal decimalPrecision As Integer,
                                             Optional ByVal numberFormatInfo As NumberFormatInfo = Nothing) As String
            <DebuggerStepThrough>
            Get
                If numberFormatInfo Is Nothing Then
                    numberFormatInfo = CultureInfo.CurrentCulture.NumberFormat
                End If
                Return Me.SizeRounded.ToString(String.Format("N{0}", decimalPrecision), numberFormatInfo)
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the <see cref="SizeUnits"/> used to round the <see cref="Filesize.Size"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The <see cref="SizeUnits"/>.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public ReadOnly Property SizeUnit As SizeUnits

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the short name of the <see cref="SizeUnits"/> used to round the <see cref="Filesize.Size"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The short name of the <see cref="SizeUnits"/>.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public ReadOnly Property SizeUnitNameShort As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the long name of the <see cref="SizeUnits"/> used to round the <see cref="Filesize.Size"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The short long of the <see cref="SizeUnits"/>.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public ReadOnly Property SizeUnitNameLong As String

#End Region

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Prevents a default instance of the <see cref="Filesize"/> class from being created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub New()
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Initializes a new instance of the <see cref="Filesize"/> class.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="size">
        ''' The filesize.
        ''' </param>
        ''' 
        ''' <param name="sizeUnit">
        ''' The unit of size.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Sub New(ByVal size As Double, ByVal sizeUnit As SizeUnits)

            Me.bytesB = Global.System.Convert.ToDouble(size * sizeUnit)

            Select Case Math.Abs(Me.bytesB)

                Case Is >= SizeUnits.Petabyte
                    Me.SizeRounded = (Me.bytesB / SizeUnits.Petabyte)
                    Me.SizeUnit = SizeUnits.Petabyte
                    Me.SizeUnitNameShort = "PB"
                    Me.SizeUnitNameLong = "PetaBytes"

                Case Is >= SizeUnits.Terabyte
                    Me.SizeRounded = (Me.bytesB / SizeUnits.Terabyte)
                    Me.SizeUnit = SizeUnits.Terabyte
                    Me.SizeUnitNameShort = "TB"
                    Me.SizeUnitNameLong = "TeraBytes"

                Case Is >= SizeUnits.Gigabyte
                    Me.SizeRounded = (Me.bytesB / SizeUnits.Gigabyte)
                    Me.SizeUnit = SizeUnits.Gigabyte
                    Me.SizeUnitNameShort = "GB"
                    Me.SizeUnitNameLong = "GigaBytes"

                Case Is >= SizeUnits.Megabyte
                    Me.SizeRounded = (Me.bytesB / SizeUnits.Megabyte)
                    Me.SizeUnit = SizeUnits.Megabyte
                    Me.SizeUnitNameShort = "MB"
                    Me.SizeUnitNameLong = "MegaBytes"

                Case Is >= SizeUnits.Kilobyte
                    Me.SizeRounded = (Me.bytesB / SizeUnits.Kilobyte)
                    Me.SizeUnit = SizeUnits.Kilobyte
                    Me.SizeUnitNameShort = "KB"
                    Me.SizeUnitNameLong = "KiloBytes"

                Case Is >= SizeUnits.Byte, Is = 0
                    Me.SizeRounded = (Me.bytesB / SizeUnits.Byte)
                    Me.SizeUnit = SizeUnits.Byte
                    Me.SizeUnitNameShort = "Bytes"
                    Me.SizeUnitNameLong = "Bytes"

            End Select

        End Sub

#End Region

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Returns a <see cref="String"/> that represents this filesize.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A <see cref="String"/> that represents this filesize.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Overrides Function ToString() As String
            Return Me.ToString(CultureInfo.InvariantCulture.NumberFormat)
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Returns a <see cref="String"/> that represents this filesize.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A <see cref="String"/> that represents this filesize.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Overloads Function ToString(ByVal provider As IFormatProvider) As String
            If Me.SizeUnit = SizeUnits.Byte Then
                Return String.Format(provider, "{0:0.##} {1}", Math.Floor(Me.SizeRounded * 100) / 100, Me.SizeUnitNameShort)
            Else
                Return String.Format(provider, "{0:0.00} {1}", Math.Floor(Me.SizeRounded * 100) / 100, Me.SizeUnitNameShort)
            End If
        End Function

#End Region

#Region " Private Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Converts the specified filesize to a different unit of size.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="size">
        ''' The filesize.
        ''' </param>
        ''' 
        ''' <param name="fromUnit">
        ''' The unit size to convert from.
        ''' </param>
        ''' 
        ''' <param name="toUnit">
        ''' The unit size to convert to.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The resulting value of the unit conversion.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="InvalidEnumArgumentException">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Private Function Convert(ByVal size As Double,
                                 ByVal fromUnit As SizeUnits,
                                 ByVal toUnit As SizeUnits) As Double

            Dim bytes As Double

            If fromUnit = SizeUnits.Byte Then
                bytes = size

            Else
                bytes = Global.System.Convert.ToDouble(size * fromUnit)

            End If

            If (toUnit < fromUnit) Then

                Select Case toUnit

                    Case SizeUnits.Byte
                        Return bytes

                    Case SizeUnits.Kilobyte
                        Return (bytes / SizeUnits.Kilobyte)

                    Case SizeUnits.Megabyte
                        Return (bytes / SizeUnits.Megabyte)

                    Case SizeUnits.Gigabyte
                        Return (bytes / SizeUnits.Gigabyte)

                    Case SizeUnits.Terabyte
                        Return (bytes / SizeUnits.Terabyte)

                    Case SizeUnits.Petabyte
                        Return (bytes / SizeUnits.Petabyte)

                    Case Else
                        Throw New InvalidEnumArgumentException(argumentName:=NameOf(toUnit), invalidValue:=CInt(toUnit),
                                                               enumClass:=GetType(SizeUnits))

                End Select

            ElseIf (toUnit > fromUnit) Then

                Select Case toUnit

                    Case SizeUnits.Byte
                        Return bytes

                    Case SizeUnits.Kilobyte
                        Return (bytes * SizeUnits.Kilobyte / SizeUnits.Kilobyte ^ 2)

                    Case SizeUnits.Megabyte
                        Return (bytes * SizeUnits.Megabyte / SizeUnits.Megabyte ^ 2)

                    Case SizeUnits.Gigabyte
                        Return (bytes * SizeUnits.Gigabyte / SizeUnits.Gigabyte ^ 2)

                    Case SizeUnits.Terabyte
                        Return (bytes * SizeUnits.Terabyte / SizeUnits.Terabyte ^ 2)

                    Case SizeUnits.Petabyte
                        Return (bytes * SizeUnits.Petabyte / SizeUnits.Petabyte ^ 2)

                    Case Else
                        Throw New InvalidEnumArgumentException(argumentName:=NameOf(toUnit), invalidValue:=CInt(toUnit),
                                                               enumClass:=GetType(SizeUnits))

                End Select

            Else ' If toUnit = fromUnit 
                Return bytes

            End If

        End Function

#End Region

    End Class

End Namespace

#End Region
