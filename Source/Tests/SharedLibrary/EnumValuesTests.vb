#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.Globalization

Imports DevCase.Core.Application.UserInterface
Imports DevCase.Core.IO

Imports Microsoft.VisualStudio.TestTools.UnitTesting

#End Region

#Region " EnumValuesTests "

Namespace EasyLink.Tests.SharedLibrary

    ''' <summary>
    ''' Unit tests that pin the numeric values of the enumerations whose values are persisted
    ''' in shortcut files, in user settings, or passed to native Windows APIs.
    ''' <para></para>
    ''' Changing any of these values would silently corrupt existing data or native calls.
    ''' </summary>
    <TestClass>
    Public NotInheritable Class EnumValuesTests

#Region " Helper Methods "

        ''' <summary>
        ''' Gets the numeric value of the enumeration member with the specified name, resolved at runtime
        ''' (so the assertions verify the compiled enumeration, not a constant folded by the compiler).
        ''' </summary>
        ''' 
        ''' <typeparam name="TEnum">
        ''' The enumeration type.
        ''' </typeparam>
        ''' 
        ''' <param name="memberName">
        ''' The name of the enumeration member.
        ''' </param>
        ''' 
        ''' <returns>
        ''' The numeric value of the member.
        ''' </returns>
        Private Shared Function ValueOf(Of TEnum As Structure)(memberName As String) As Long

            Dim member As Object = [Enum].Parse(GetType(TEnum), memberName, ignoreCase:=False)
            Return Convert.ToInt64(member, CultureInfo.InvariantCulture)

        End Function

#End Region

#Region " Tests "

        ''' <summary>
        ''' Verifies that <see cref="ShortcutWindowState"/> matches the native SW_* "show command" values
        ''' used by IShellLink::SetShowCmd.
        ''' </summary>
        <TestMethod>
        Public Sub ShortcutWindowState_MatchesNativeShowCommands()

            Const SW_SHOWNORMAL As Long = 1
            Const SW_SHOWMAXIMIZED As Long = 3
            Const SW_SHOWMINNOACTIVE As Long = 7

            Assert.AreEqual(SW_SHOWNORMAL, EnumValuesTests.ValueOf(Of ShortcutWindowState)("Normal"))
            Assert.AreEqual(SW_SHOWMAXIMIZED, EnumValuesTests.ValueOf(Of ShortcutWindowState)("Maximized"))
            Assert.AreEqual(SW_SHOWMINNOACTIVE, EnumValuesTests.ValueOf(Of ShortcutWindowState)("Minimized"))
            Assert.AreEqual(3, [Enum].GetValues(GetType(ShortcutWindowState)).Length)

        End Sub

        ''' <summary>
        ''' Verifies that <see cref="ShortcutHotkeyModifier"/> matches the native HOTKEYF_* values
        ''' used by IShellLink::SetHotkey.
        ''' </summary>
        <TestMethod>
        Public Sub ShortcutHotkeyModifier_MatchesNativeHotkeyFlags()

            Const HOTKEYF_SHIFT As Long = &H1
            Const HOTKEYF_CONTROL As Long = &H2
            Const HOTKEYF_ALT As Long = &H4

            Assert.AreEqual(0L, EnumValuesTests.ValueOf(Of ShortcutHotkeyModifier)("None"))
            Assert.AreEqual(HOTKEYF_SHIFT, EnumValuesTests.ValueOf(Of ShortcutHotkeyModifier)("Shift"))
            Assert.AreEqual(HOTKEYF_CONTROL, EnumValuesTests.ValueOf(Of ShortcutHotkeyModifier)("Control"))
            Assert.AreEqual(HOTKEYF_ALT, EnumValuesTests.ValueOf(Of ShortcutHotkeyModifier)("Alt"))
            Assert.AreEqual(GetType(Short), [Enum].GetUnderlyingType(GetType(ShortcutHotkeyModifier)))
            Assert.IsTrue(Attribute.IsDefined(GetType(ShortcutHotkeyModifier), GetType(FlagsAttribute)))

        End Sub

        ''' <summary>
        ''' Verifies the values of <see cref="ShortcutTargetExecutionMode"/>.
        ''' </summary>
        <TestMethod>
        Public Sub ShortcutTargetExecutionMode_HasExpectedValues()

            Assert.AreEqual(0L, EnumValuesTests.ValueOf(Of ShortcutTargetExecutionMode)("Standard"))
            Assert.AreEqual(1L, EnumValuesTests.ValueOf(Of ShortcutTargetExecutionMode)("RelativePathViaExplorer"))
            Assert.AreEqual(2L, EnumValuesTests.ValueOf(Of ShortcutTargetExecutionMode)("RelativePathWithArgumentsViaCommandPrompt"))
            Assert.AreEqual(3, [Enum].GetValues(GetType(ShortcutTargetExecutionMode)).Length)

        End Sub

        ''' <summary>
        ''' Verifies the values of <see cref="VisualTheme"/>, which are persisted in the user settings.
        ''' </summary>
        <TestMethod>
        Public Sub VisualTheme_HasExpectedValues()

            Assert.AreEqual(0L, EnumValuesTests.ValueOf(Of VisualTheme)("Default"))
            Assert.AreEqual(1L, EnumValuesTests.ValueOf(Of VisualTheme)("VisualStudioDark"))

        End Sub

        ''' <summary>
        ''' Verifies the values of the CLI <see cref="CommandLineAction"/> enumeration,
        ''' where <see cref="CommandLineAction.None"/> must be the default value.
        ''' </summary>
        <TestMethod>
        Public Sub CommandLineAction_NoneIsDefault()

            Assert.AreEqual(0L, EnumValuesTests.ValueOf(Of CommandLineAction)("None"))
            Assert.AreEqual(5, [Enum].GetValues(GetType(CommandLineAction)).Length)

        End Sub

#End Region

    End Class

End Namespace

#End Region
