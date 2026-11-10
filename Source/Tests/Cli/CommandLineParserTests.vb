#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.Collections.Generic
Imports System.Windows.Forms

Imports DevCase.Core.IO

Imports EasyLink.Tests.Helpers

Imports Microsoft.VisualStudio.TestTools.UnitTesting

#End Region

#Region " CommandLineParserTests "

Namespace EasyLink.Tests.Cli

    ''' <summary>
    ''' Unit tests for the <see cref="CommandLineParser"/> class of the CLI application.
    ''' <para></para>
    ''' These tests are pure: they do not touch the file system or the Windows Shell.
    ''' </summary>
    <TestClass>
    Public NotInheritable Class CommandLineParserTests

#Region " IsCliMode "

        ''' <summary>
        ''' Verifies that an empty argument list is not considered CLI mode.
        ''' </summary>
        <TestMethod>
        Public Sub IsCliMode_EmptyArguments_ReturnsFalse()

            Dim result As Boolean = CommandLineParser.IsCliMode(New String() {})

            Assert.IsFalse(result)

        End Sub

        ''' <summary>
        ''' Verifies that a single file path (the GUI "open with" scenario) is not considered CLI mode.
        ''' </summary>
        <TestMethod>
        Public Sub IsCliMode_OnlyFilePath_ReturnsFalse()

            Dim result As Boolean = CommandLineParser.IsCliMode(New String() {"C:\Shortcut.lnk"})

            Assert.IsFalse(result)

        End Sub

        ''' <summary>
        ''' Verifies that any argument starting with a slash enables CLI mode, wherever it appears.
        ''' </summary>
        <TestMethod>
        <DataRow("/Read")>
        <DataRow("/?")>
        <DataRow("/UnknownSwitch")>
        Public Sub IsCliMode_ArgumentsWithSwitch_ReturnsTrue(switchArgument As String)

            Dim result As Boolean = CommandLineParser.IsCliMode(New String() {"C:\Shortcut.lnk", switchArgument})

            Assert.IsTrue(result)

        End Sub

#End Region

#Region " Parse - Defaults "

        ''' <summary>
        ''' Verifies that parsing an empty argument list returns the default option values,
        ''' where every optional property is <see langword="Nothing"/> (meaning "not specified").
        ''' </summary>
        <TestMethod>
        Public Sub Parse_EmptyArguments_ReturnsDefaultOptions()

            Dim options As CommandLineOptions = CommandLineParser.Parse(New String() {})

            Assert.AreEqual(CommandLineAction.None, options.Action)
            Assert.AreEqual(String.Empty, options.FilePath)
            Assert.IsNull(options.Target)
            Assert.IsNull(options.TargetArguments)
            Assert.IsNull(options.Description)
            Assert.IsNull(options.Icon)
            Assert.IsFalse(options.IconIndex.HasValue)
            Assert.IsFalse(options.Hotkey.HasValue)
            Assert.IsFalse(options.WindowState.HasValue)
            Assert.IsNull(options.WorkingDirectory)
            Assert.IsNull(options.AppId)
            Assert.IsFalse(options.Force)

        End Sub

#End Region

#Region " Parse - Actions "

        ''' <summary>
        ''' Verifies that each action switch is recognized, case-insensitively.
        ''' </summary>
        <TestMethod>
        <DataRow("/Create", "Create")>
        <DataRow("/create", "Create")>
        <DataRow("/CREATE", "Create")>
        <DataRow("/Modify", "Modify")>
        <DataRow("/mOdIfY", "Modify")>
        <DataRow("/Read", "Read")>
        <DataRow("/read", "Read")>
        <DataRow("/Help", "Help")>
        <DataRow("/help", "Help")>
        <DataRow("/?", "Help")>
        Public Sub Parse_ActionSwitch_SetsAction(switchArgument As String, expectedActionName As String)

            ' The expected action is received by name because CommandLineAction is a Friend type,
            ' which cannot appear in the signature of a public test method.
            Dim expectedAction As CommandLineAction =
                CType([Enum].Parse(GetType(CommandLineAction), expectedActionName), CommandLineAction)

            Dim options As CommandLineOptions = CommandLineParser.Parse(New String() {switchArgument})

            Assert.AreEqual(expectedAction, options.Action)

        End Sub

        ''' <summary>
        ''' Verifies that specifying more than one action throws an <see cref="ArgumentException"/>.
        ''' </summary>
        <TestMethod>
        <DataRow("/Create", "/Read")>
        <DataRow("/Modify", "/Create")>
        <DataRow("/Read", "/Help")>
        <DataRow("/Read", "/Read")>
        Public Sub Parse_MultipleActions_ThrowsArgumentException(firstAction As String, secondAction As String)

            Dim ex As ArgumentException = Assert.ThrowsExactly(Of ArgumentException)(
                Sub() CommandLineParser.Parse(New String() {firstAction, secondAction}))

            TextAssert.Contains(ex.Message, "Multiple actions specified")

        End Sub

        ''' <summary>
        ''' Verifies that an unknown switch throws an <see cref="ArgumentException"/> that names the switch.
        ''' </summary>
        <TestMethod>
        <DataRow("/Delete")>
        <DataRow("/h")>
        <DataRow("/")>
        <DataRow("/:value")>
        Public Sub Parse_UnknownSwitch_ThrowsArgumentException(switchArgument As String)

            Dim ex As ArgumentException = Assert.ThrowsExactly(Of ArgumentException)(
                Sub() CommandLineParser.Parse(New String() {"/Read", switchArgument}))

            TextAssert.Contains(ex.Message, "Unknown switch")

        End Sub

#End Region

#Region " Parse - File Path "

        ''' <summary>
        ''' Verifies that the /File switch sets the file path.
        ''' </summary>
        <TestMethod>
        Public Sub Parse_FileSwitch_SetsFilePath()

            Dim options As CommandLineOptions = CommandLineParser.Parse(New String() {"/Read", "/File:C:\Dir\My Shortcut.lnk"})

            Assert.AreEqual("C:\Dir\My Shortcut.lnk", options.FilePath)

        End Sub

        ''' <summary>
        ''' Verifies that a non-switch argument is taken as the file path when /File was not specified.
        ''' </summary>
        <TestMethod>
        Public Sub Parse_PositionalArgument_SetsFilePath()

            Dim options As CommandLineOptions = CommandLineParser.Parse(New String() {"/Read", "C:\Positional.lnk"})

            Assert.AreEqual("C:\Positional.lnk", options.FilePath)

        End Sub

        ''' <summary>
        ''' Verifies that only the first positional argument is used as the file path.
        ''' </summary>
        <TestMethod>
        Public Sub Parse_MultiplePositionalArguments_FirstOneWins()

            Dim options As CommandLineOptions = CommandLineParser.Parse(New String() {"C:\First.lnk", "C:\Second.lnk", "/Read"})

            Assert.AreEqual("C:\First.lnk", options.FilePath)

        End Sub

        ''' <summary>
        ''' Verifies that a positional argument does not override a previously specified /File value.
        ''' </summary>
        <TestMethod>
        Public Sub Parse_PositionalAfterFileSwitch_DoesNotOverrideFilePath()

            Dim options As CommandLineOptions = CommandLineParser.Parse(New String() {"/File:C:\Explicit.lnk", "C:\Positional.lnk"})

            Assert.AreEqual("C:\Explicit.lnk", options.FilePath)

        End Sub

        ''' <summary>
        ''' Verifies that a /File switch specified after a positional argument overrides it.
        ''' </summary>
        <TestMethod>
        Public Sub Parse_FileSwitchAfterPositional_OverridesFilePath()

            Dim options As CommandLineOptions = CommandLineParser.Parse(New String() {"C:\Positional.lnk", "/File:C:\Explicit.lnk"})

            Assert.AreEqual("C:\Explicit.lnk", options.FilePath)

        End Sub

#End Region

#Region " Parse - String Properties "

        ''' <summary>
        ''' Verifies that every string property switch is mapped to its option, case-insensitively,
        ''' and that only the first colon separates the switch name from its value
        ''' (so values like drive-rooted paths are preserved).
        ''' </summary>
        <TestMethod>
        Public Sub Parse_StringPropertySwitches_SetProperties()

            Dim args As String() = {
                "/Modify",
                "/File:C:\Shortcut.lnk",
                "/Target:C:\Windows\notepad.exe",
                "/ARGS:--mode:fast ""C:\file.txt""",
                "/desc:Description: with colons",
                "/Icon:C:\Windows\System32\shell32.dll",
                "/WorkDir:C:\Windows",
                "/AppId:Company.Product.App"
            }

            Dim options As CommandLineOptions = CommandLineParser.Parse(args)

            Assert.AreEqual(CommandLineAction.Modify, options.Action)
            Assert.AreEqual("C:\Shortcut.lnk", options.FilePath)
            Assert.AreEqual("C:\Windows\notepad.exe", options.Target)
            Assert.AreEqual("--mode:fast ""C:\file.txt""", options.TargetArguments)
            Assert.AreEqual("Description: with colons", options.Description)
            Assert.AreEqual("C:\Windows\System32\shell32.dll", options.Icon)
            Assert.AreEqual("C:\Windows", options.WorkingDirectory)
            Assert.AreEqual("Company.Product.App", options.AppId)

        End Sub

        ''' <summary>
        ''' Verifies that a switch written with a colon but no value sets an empty string,
        ''' which is distinct from <see langword="Nothing"/> ("not specified").
        ''' This allows clearing a property with the /Modify action.
        ''' </summary>
        <TestMethod>
        Public Sub Parse_SwitchWithEmptyValue_SetsEmptyString()

            Dim options As CommandLineOptions = CommandLineParser.Parse(New String() {"/Modify", "/Desc:", "/Args:"})

            Assert.AreEqual(String.Empty, options.Description)
            Assert.AreEqual(String.Empty, options.TargetArguments)

        End Sub

        ''' <summary>
        ''' Verifies that every switch that requires a value throws an <see cref="ArgumentException"/>
        ''' when written without the colon and value.
        ''' </summary>
        <TestMethod>
        <DataRow("/File")>
        <DataRow("/Target")>
        <DataRow("/Args")>
        <DataRow("/Desc")>
        <DataRow("/Icon")>
        <DataRow("/IconIndex")>
        <DataRow("/Hotkey")>
        <DataRow("/WindowState")>
        <DataRow("/WorkDir")>
        <DataRow("/AppId")>
        Public Sub Parse_ValueSwitchWithoutValue_ThrowsArgumentException(switchArgument As String)

            Dim ex As ArgumentException = Assert.ThrowsExactly(Of ArgumentException)(
                Sub() CommandLineParser.Parse(New String() {"/Create", switchArgument}))

            TextAssert.Contains(ex.Message, "requires a value")

        End Sub

#End Region

#Region " Parse - Force "

        ''' <summary>
        ''' Verifies that the /Force flag switch is recognized.
        ''' </summary>
        <TestMethod>
        <DataRow("/Force")>
        <DataRow("/force")>
        Public Sub Parse_ForceSwitch_SetsForce(switchArgument As String)

            Dim options As CommandLineOptions = CommandLineParser.Parse(New String() {"/Create", switchArgument})

            Assert.IsTrue(options.Force)

        End Sub

#End Region

#Region " Parse - IconIndex "

        ''' <summary>
        ''' Verifies that valid integer icon indexes (including negative resource IDs) are parsed.
        ''' </summary>
        <TestMethod>
        <DataRow("0", 0)>
        <DataRow("5", 5)>
        <DataRow("-3", -3)>
        <DataRow("2147483647", Integer.MaxValue)>
        Public Sub Parse_IconIndexValid_SetsIconIndex(value As String, expected As Integer)

            Dim options As CommandLineOptions = CommandLineParser.Parse(New String() {"/Create", $"/IconIndex:{value}"})

            Assert.IsTrue(options.IconIndex.HasValue)
            Assert.AreEqual(expected, options.IconIndex.Value)

        End Sub

        ''' <summary>
        ''' Verifies that non-integer icon indexes throw an <see cref="ArgumentException"/>.
        ''' </summary>
        <TestMethod>
        <DataRow("abc")>
        <DataRow("1.5")>
        <DataRow("")>
        <DataRow("99999999999")>
        Public Sub Parse_IconIndexInvalid_ThrowsArgumentException(value As String)

            Dim ex As ArgumentException = Assert.ThrowsExactly(Of ArgumentException)(
                Sub() CommandLineParser.Parse(New String() {"/Create", $"/IconIndex:{value}"}))

            TextAssert.Contains(ex.Message, "Invalid /IconIndex value")

        End Sub

#End Region

#Region " Parse - Hotkey "

        ''' <summary>
        ''' Verifies the supported hotkey formats: "+" separated names, "," separated names,
        ''' case-insensitive names, numeric values and "None".
        ''' </summary>
        <TestMethod>
        <DataRow("Control+Alt+F1", Keys.Control Or Keys.Alt Or Keys.F1)>
        <DataRow("control+alt+f1", Keys.Control Or Keys.Alt Or Keys.F1)>
        <DataRow("Control + Shift + A", Keys.Control Or Keys.Shift Or Keys.A)>
        <DataRow("Shift, Alt, F12", Keys.Shift Or Keys.Alt Or Keys.F12)>
        <DataRow("F5", Keys.F5)>
        <DataRow("None", Keys.None)>
        <DataRow("131184", Keys.Control Or Keys.F1)>
        <DataRow("0", Keys.None)>
        Public Sub Parse_HotkeyValid_SetsHotkey(value As String, expected As Keys)

            Dim options As CommandLineOptions = CommandLineParser.Parse(New String() {"/Create", $"/Hotkey:{value}"})

            Assert.IsTrue(options.Hotkey.HasValue)
            Assert.AreEqual(expected, options.Hotkey.Value)

        End Sub

        ''' <summary>
        ''' Verifies that invalid hotkey names throw an <see cref="ArgumentException"/>.
        ''' </summary>
        <TestMethod>
        <DataRow("NotAKey")>
        <DataRow("Control+NotAKey")>
        <DataRow("")>
        Public Sub Parse_HotkeyInvalid_ThrowsArgumentException(value As String)

            Dim ex As ArgumentException = Assert.ThrowsExactly(Of ArgumentException)(
                Sub() CommandLineParser.Parse(New String() {"/Create", $"/Hotkey:{value}"}))

            TextAssert.Contains(ex.Message, "Invalid /Hotkey value")

        End Sub

#End Region

#Region " Parse - WindowState "

        ''' <summary>
        ''' Verifies the supported window state values: names (case-insensitive),
        ''' the Min/Max aliases, and the numeric values of the defined members.
        ''' </summary>
        <TestMethod>
        <DataRow("Normal", ShortcutWindowState.Normal)>
        <DataRow("normal", ShortcutWindowState.Normal)>
        <DataRow("Minimized", ShortcutWindowState.Minimized)>
        <DataRow("MINIMIZED", ShortcutWindowState.Minimized)>
        <DataRow("Maximized", ShortcutWindowState.Maximized)>
        <DataRow("Min", ShortcutWindowState.Minimized)>
        <DataRow("min", ShortcutWindowState.Minimized)>
        <DataRow("Max", ShortcutWindowState.Maximized)>
        <DataRow("MAX", ShortcutWindowState.Maximized)>
        <DataRow(" Max ", ShortcutWindowState.Maximized)>
        <DataRow("1", ShortcutWindowState.Normal)>
        <DataRow("3", ShortcutWindowState.Maximized)>
        <DataRow("7", ShortcutWindowState.Minimized)>
        Public Sub Parse_WindowStateValid_SetsWindowState(value As String, expected As ShortcutWindowState)

            Dim options As CommandLineOptions = CommandLineParser.Parse(New String() {"/Create", $"/WindowState:{value}"})

            Assert.IsTrue(options.WindowState.HasValue)
            Assert.AreEqual(expected, options.WindowState.Value)

        End Sub

        ''' <summary>
        ''' Verifies that undefined numeric values, comma-separated combinations and unknown names
        ''' are rejected with an <see cref="ArgumentException"/>, instead of producing an
        ''' undefined <see cref="ShortcutWindowState"/> value (regression test).
        ''' </summary>
        <TestMethod>
        <DataRow("0")>
        <DataRow("2")>
        <DataRow("5")>
        <DataRow("-1")>
        <DataRow("Normal, Maximized")>
        <DataRow("Normal,Minimized")>
        <DataRow("Hidden")>
        <DataRow("")>
        Public Sub Parse_WindowStateInvalid_ThrowsArgumentException(value As String)

            Dim ex As ArgumentException = Assert.ThrowsExactly(Of ArgumentException)(
                Sub() CommandLineParser.Parse(New String() {"/Create", $"/WindowState:{value}"}))

            TextAssert.Contains(ex.Message, "Invalid /WindowState value")

        End Sub

#End Region

#Region " Parse - Full Command Lines "

        ''' <summary>
        ''' Verifies a realistic full /Create command line, with switches in arbitrary order.
        ''' </summary>
        <TestMethod>
        Public Sub Parse_FullCreateCommandLine_SetsAllOptions()

            Dim args As New List(Of String) From {
                "/Target:C:\Windows\notepad.exe",
                "/Create",
                "/IconIndex:2",
                "/Hotkey:Control+Alt+N",
                "/WindowState:Max",
                "/File:C:\Notepad.lnk",
                "/Force"
            }

            Dim options As CommandLineOptions = CommandLineParser.Parse(args)

            Assert.AreEqual(CommandLineAction.Create, options.Action)
            Assert.AreEqual("C:\Notepad.lnk", options.FilePath)
            Assert.AreEqual("C:\Windows\notepad.exe", options.Target)
            Assert.AreEqual(2, options.IconIndex.Value)
            Assert.AreEqual(Keys.Control Or Keys.Alt Or Keys.N, options.Hotkey.Value)
            Assert.AreEqual(ShortcutWindowState.Maximized, options.WindowState.Value)
            Assert.IsTrue(options.Force)

            ' Properties not specified must remain "not specified".
            Assert.IsNull(options.Description)
            Assert.IsNull(options.TargetArguments)
            Assert.IsNull(options.Icon)
            Assert.IsNull(options.WorkingDirectory)
            Assert.IsNull(options.AppId)

        End Sub

        ''' <summary>
        ''' Verifies that when the same property switch is specified twice, the last value wins.
        ''' </summary>
        <TestMethod>
        Public Sub Parse_RepeatedPropertySwitch_LastValueWins()

            Dim options As CommandLineOptions = CommandLineParser.Parse(New String() {"/Modify", "/Desc:First", "/Desc:Second"})

            Assert.AreEqual("Second", options.Description)

        End Sub

#End Region

    End Class

End Namespace

#End Region
