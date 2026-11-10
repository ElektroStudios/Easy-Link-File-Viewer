#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.ComponentModel.Design.Serialization
Imports System.Globalization

Imports DevCase.Core.Design

Imports EasyLink.Tests.Helpers

Imports Microsoft.VisualStudio.TestTools.UnitTesting

#End Region

#Region " FileSizeConverterTests "

Namespace EasyLink.Tests.SharedLibrary

    ''' <summary>
    ''' Unit tests for the <see cref="FileSizeConverter"/> class, used by the property grid
    ''' to display the shortcut file length the same way as Windows Explorer.
    ''' </summary>
    <TestClass>
    Public NotInheritable Class FileSizeConverterTests

        ''' <summary>
        ''' Verifies that numeric sizes are formatted with the Windows size unit
        ''' (the decimal separator depends on the user locale, so only the unit is asserted).
        ''' </summary>
        <TestMethod>
        <DataRow(2048L, "KB")>
        <DataRow(5L * 1024L * 1024L, "MB")>
        <DataRow(3L * 1024L * 1024L * 1024L, "GB")>
        Public Sub ConvertTo_NumericValue_ReturnsWindowsFormattedSize(value As Long, expectedUnit As String)

            Dim converter As New FileSizeConverter()

            Dim result As String = CStr(converter.ConvertTo(Nothing, CultureInfo.InvariantCulture, value, GetType(String)))

            Assert.IsFalse(String.IsNullOrWhiteSpace(result))
            Assert.IsTrue(result.EndsWith(expectedUnit, StringComparison.Ordinal), $"Unexpected formatted size: ""{result}""")

        End Sub

        ''' <summary>
        ''' Verifies that non-numeric values are returned unchanged.
        ''' </summary>
        <TestMethod>
        Public Sub ConvertTo_NonNumericValue_ReturnsOriginalValue()

            Dim converter As New FileSizeConverter()

            Dim result As Object = converter.ConvertTo(Nothing, CultureInfo.InvariantCulture, "not a number", GetType(String))

            Assert.AreEqual("not a number", CStr(result))

        End Sub

        ''' <summary>
        ''' Verifies that a <see langword="Nothing"/> destination type is rejected.
        ''' </summary>
        <TestMethod>
        Public Sub ConvertTo_NullDestinationType_ThrowsArgumentNullException()

            Dim converter As New FileSizeConverter()

            Assert.ThrowsExactly(Of ArgumentNullException)(
                Sub() converter.ConvertTo(Nothing, CultureInfo.InvariantCulture, 1024L, Nothing))

        End Sub

        ''' <summary>
        ''' Verifies that the converter is display-only: it cannot convert from other types.
        ''' </summary>
        <TestMethod>
        Public Sub CanConvertFrom_AnyType_ReturnsFalse()

            Dim converter As New FileSizeConverter()

            Assert.IsFalse(converter.CanConvertFrom(Nothing, GetType(String)))
            Assert.IsFalse(converter.CanConvertFrom(Nothing, GetType(Long)))

        End Sub

        ''' <summary>
        ''' Verifies that the converter supports conversion to strings and instance descriptors.
        ''' </summary>
        <TestMethod>
        Public Sub CanConvertTo_StringAndInstanceDescriptor_ReturnsTrue()

            Dim converter As New FileSizeConverter()

            Assert.IsTrue(converter.CanConvertTo(Nothing, GetType(String)))
            Assert.IsTrue(converter.CanConvertTo(Nothing, GetType(InstanceDescriptor)))

        End Sub

    End Class

End Namespace

#End Region
