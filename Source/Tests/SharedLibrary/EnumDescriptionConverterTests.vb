#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.ComponentModel
Imports System.Globalization

Imports DevCase.Core.IO
Imports DevCase.Runtime.TypeConverters

Imports Microsoft.VisualStudio.TestTools.UnitTesting

#End Region

#Region " Test Enumerations "

Namespace EasyLink.Tests.SharedLibrary

    ''' <summary>
    ''' A non-flags enumeration used to test <see cref="EnumDescriptionConverter"/>.
    ''' </summary>
    Public Enum TestColor As Integer

        ''' <summary>
        ''' A member with a description.
        ''' </summary>
        <System.ComponentModel.Description("Pure Red")>
        Red = 0

        ''' <summary>
        ''' A member without a description.
        ''' </summary>
        Green = 1

    End Enum

    ''' <summary>
    ''' A flags enumeration used to test <see cref="EnumDescriptionConverter"/>.
    ''' </summary>
    <Flags>
    Public Enum TestFlags As Integer

        ''' <summary>
        ''' No flags.
        ''' </summary>
        None = 0

        ''' <summary>
        ''' A flag with a description.
        ''' </summary>
        <System.ComponentModel.Description("First Flag")>
        First = 1 << 0

        ''' <summary>
        ''' Another flag with a description.
        ''' </summary>
        <System.ComponentModel.Description("Second Flag")>
        Second = 1 << 1

        ''' <summary>
        ''' A flag whose description contains a comma, which the converter must replace by the member name.
        ''' </summary>
        <System.ComponentModel.Description("Third, with comma")>
        Third = 1 << 2

    End Enum

End Namespace

#End Region

#Region " EnumDescriptionConverterTests "

Namespace EasyLink.Tests.SharedLibrary

    ''' <summary>
    ''' Unit tests for the <see cref="EnumDescriptionConverter"/> class.
    ''' </summary>
    <TestClass>
    Public NotInheritable Class EnumDescriptionConverterTests

#Region " Non-Flags Enumerations "

        ''' <summary>
        ''' Verifies that a member is converted to its description, or to its name when it has no description.
        ''' </summary>
        <TestMethod>
        <DataRow(TestColor.Red, "Pure Red")>
        <DataRow(TestColor.Green, "Green")>
        Public Sub ConvertTo_NonFlags_ReturnsDescriptionOrName(value As TestColor, expected As String)

            Dim converter As New EnumDescriptionConverter(GetType(TestColor))

            Dim result As Object = converter.ConvertTo(Nothing, CultureInfo.InvariantCulture, value, GetType(String))

            Assert.AreEqual(expected, result)

        End Sub

        ''' <summary>
        ''' Verifies that a description or a member name is converted back to the member.
        ''' </summary>
        <TestMethod>
        <DataRow("Pure Red", TestColor.Red)>
        <DataRow("Red", TestColor.Red)>
        <DataRow("Green", TestColor.Green)>
        Public Sub ConvertFrom_NonFlags_ReturnsMember(text As String, expected As TestColor)

            Dim converter As New EnumDescriptionConverter(GetType(TestColor))

            Dim result As Object = converter.ConvertFrom(Nothing, CultureInfo.InvariantCulture, text)

            Assert.AreEqual(expected, CType(result, TestColor))

        End Sub

        ''' <summary>
        ''' Verifies that an unknown text throws an <see cref="ArgumentException"/>.
        ''' </summary>
        <TestMethod>
        Public Sub ConvertFrom_NonFlagsUnknownText_ThrowsArgumentException()

            Dim converter As New EnumDescriptionConverter(GetType(TestColor))

            Assert.ThrowsExactly(Of ArgumentException)(
                Sub() converter.ConvertFrom(Nothing, CultureInfo.InvariantCulture, "Blue"))

        End Sub

        ''' <summary>
        ''' Verifies the standard values support, which lists every member and allows free text.
        ''' </summary>
        <TestMethod>
        Public Sub StandardValues_NonFlags_ListsAllMembers()

            Dim converter As New EnumDescriptionConverter(GetType(TestColor))

            Assert.IsTrue(converter.GetStandardValuesSupported(Nothing))
            Assert.IsFalse(converter.GetStandardValuesExclusive(Nothing))
            Assert.AreEqual(2, converter.GetStandardValues(Nothing).Count)

        End Sub

        ''' <summary>
        ''' Verifies that the converter supports conversions from and to strings.
        ''' </summary>
        <TestMethod>
        Public Sub CanConvert_String_ReturnsTrue()

            Dim converter As New EnumDescriptionConverter(GetType(TestColor))

            Assert.IsTrue(converter.CanConvertTo(Nothing, GetType(String)))
            Assert.IsTrue(converter.CanConvertFrom(Nothing, GetType(String)))

        End Sub

#End Region

#Region " Flags Enumerations "

        ''' <summary>
        ''' Verifies the conversion of flag combinations to text.
        ''' </summary>
        <TestMethod>
        <DataRow(TestFlags.None, "None")>
        <DataRow(TestFlags.First, "First Flag")>
        <DataRow(TestFlags.First Or TestFlags.Second, "First Flag, Second Flag")>
        <DataRow(TestFlags.Third, "Third")>
        <DataRow(TestFlags.First Or TestFlags.Third, "First Flag, Third")>
        <DataRow(TestFlags.First Or TestFlags.Second Or TestFlags.Third, "First Flag, Second Flag, Third")>
        Public Sub ConvertTo_Flags_ReturnsJoinedDescriptions(value As TestFlags, expected As String)

            Dim converter As New EnumDescriptionConverter(GetType(TestFlags))

            Dim result As Object = converter.ConvertTo(Nothing, CultureInfo.InvariantCulture, value, GetType(String))

            Assert.AreEqual(expected, result)

        End Sub

        ''' <summary>
        ''' Verifies the conversion of text to flag combinations, using descriptions, names (case-insensitive)
        ''' and numeric values.
        ''' </summary>
        <TestMethod>
        <DataRow("First Flag", TestFlags.First)>
        <DataRow("first flag, SECOND", TestFlags.First Or TestFlags.Second)>
        <DataRow("Third", TestFlags.Third)>
        <DataRow("None", TestFlags.None)>
        <DataRow("3", TestFlags.First Or TestFlags.Second)>
        <DataRow(" 7 ", TestFlags.First Or TestFlags.Second Or TestFlags.Third)>
        <DataRow("0", TestFlags.None)>
        Public Sub ConvertFrom_Flags_ReturnsCombination(text As String, expected As TestFlags)

            Dim converter As New EnumDescriptionConverter(GetType(TestFlags))

            Dim result As Object = converter.ConvertFrom(Nothing, CultureInfo.InvariantCulture, text)

            Assert.AreEqual(expected, CType(result, TestFlags))

        End Sub

        ''' <summary>
        ''' Verifies that unknown names and numeric values outside the defined flags are rejected.
        ''' </summary>
        <TestMethod>
        <DataRow("Fourth")>
        <DataRow("First Flag, Unknown")>
        <DataRow("8")>
        <DataRow("15")>
        Public Sub ConvertFrom_FlagsInvalidText_ThrowsArgumentException(text As String)

            Dim converter As New EnumDescriptionConverter(GetType(TestFlags))

            Assert.ThrowsExactly(Of ArgumentException)(
                Sub() converter.ConvertFrom(Nothing, CultureInfo.InvariantCulture, text))

        End Sub

        ''' <summary>
        ''' Verifies that converting every flag combination to text and back is lossless.
        ''' </summary>
        <TestMethod>
        Public Sub Flags_RoundTrip_AllCombinations()

            Dim converter As New EnumDescriptionConverter(GetType(TestFlags))

            For rawValue As Integer = 0 To 7
                Dim value As TestFlags = CType(rawValue, TestFlags)
                Dim text As String = CStr(converter.ConvertTo(Nothing, CultureInfo.InvariantCulture, value, GetType(String)))
                Dim result As TestFlags = CType(converter.ConvertFrom(Nothing, CultureInfo.InvariantCulture, text), TestFlags)
                Assert.AreEqual(value, result, $"Round-trip failed for value {rawValue} (text: ""{text}"").")
            Next rawValue

        End Sub

#End Region

#Region " ShortcutTargetExecutionMode "

        ''' <summary>
        ''' Verifies that <see cref="ShortcutTargetExecutionMode"/> is decorated with <see cref="EnumDescriptionConverter"/>,
        ''' so the property grid displays the friendly descriptions.
        ''' </summary>
        <TestMethod>
        Public Sub ShortcutTargetExecutionMode_UsesEnumDescriptionConverter()

            Dim converter As TypeConverter = TypeDescriptor.GetConverter(GetType(ShortcutTargetExecutionMode))

            Assert.IsInstanceOfType(Of EnumDescriptionConverter)(converter)

        End Sub

        ''' <summary>
        ''' Verifies the friendly descriptions of <see cref="ShortcutTargetExecutionMode"/>.
        ''' </summary>
        <TestMethod>
        <DataRow(ShortcutTargetExecutionMode.Standard, "Standard Mode")>
        <DataRow(ShortcutTargetExecutionMode.RelativePathViaExplorer, "Relative Path via Windows Explorer")>
        <DataRow(ShortcutTargetExecutionMode.RelativePathWithArgumentsViaCommandPrompt, "Relative Path with Arguments via Command Prompt")>
        Public Sub ShortcutTargetExecutionMode_DescriptionsRoundTrip(value As ShortcutTargetExecutionMode, expectedText As String)

            Dim converter As TypeConverter = TypeDescriptor.GetConverter(GetType(ShortcutTargetExecutionMode))

            Dim text As String = converter.ConvertToString(Nothing, CultureInfo.InvariantCulture, value)
            Dim result As ShortcutTargetExecutionMode =
                CType(converter.ConvertFromString(Nothing, CultureInfo.InvariantCulture, text), ShortcutTargetExecutionMode)

            Assert.AreEqual(expectedText, text)
            Assert.AreEqual(value, result)

        End Sub

#End Region

    End Class

End Namespace

#End Region
