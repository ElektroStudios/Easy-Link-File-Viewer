#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.IO
Imports System.Windows.Forms

Imports DevCase.Core.IO

Imports EasyLink.Tests.Helpers

Imports Microsoft.VisualStudio.TestTools.UnitTesting

#End Region

#Region " CommandLineProcessorTests "

Namespace EasyLink.Tests.Cli

    ''' <summary>
    ''' Unit and integration tests for the <see cref="CommandLineProcessor"/> class of the CLI application.
    ''' <para></para>
    ''' The integration tests create, modify and read real shortcut files inside a temporary directory
    ''' through the Windows Shell COM objects, so they run on a single-threaded apartment (STA) thread,
    ''' exactly like the CLI application entry point.
    ''' </summary>
    <STATestClass>
    Public NotInheritable Class CommandLineProcessorTests

#Region " Private Fields "

        ''' <summary>
        ''' Exit code for success.
        ''' </summary>
        Private Const ExitSuccess As Integer = 0

        ''' <summary>
        ''' Exit code for invalid arguments or validation errors.
        ''' </summary>
        Private Const ExitArgumentError As Integer = 1

        ''' <summary>
        ''' Exit code for runtime errors.
        ''' </summary>
        Private Const ExitRuntimeError As Integer = 2

        ''' <summary>
        ''' The temporary directory of the current test.
        ''' </summary>
        Private tempDir As TempDirectory

#End Region

#Region " Test Initialization / Cleanup "

        ''' <summary>
        ''' Creates a fresh temporary directory before each test.
        ''' </summary>
        <TestInitialize>
        Public Sub TestInitialize()
            Me.tempDir = New TempDirectory()
        End Sub

        ''' <summary>
        ''' Deletes the temporary directory after each test.
        ''' </summary>
        <TestCleanup>
        Public Sub TestCleanup()
            Me.tempDir.Dispose()
        End Sub

#End Region

#Region " Helper Methods "

        ''' <summary>
        ''' Runs the CLI processor with the specified arguments, capturing the console output.
        ''' </summary>
        '''
        ''' <param name="output">
        ''' Receives the text written to the standard output.
        ''' </param>
        '''
        ''' <param name="errorOutput">
        ''' Receives the text written to the standard error.
        ''' </param>
        '''
        ''' <param name="args">
        ''' The command-line arguments.
        ''' </param>
        '''
        ''' <returns>
        ''' The exit code returned by the processor.
        ''' </returns>
        Private Shared Function RunCli(ByRef output As String, ByRef errorOutput As String, ParamArray args As String()) As Integer

            Dim processor As New CommandLineProcessor()
            Using capture As New ConsoleCapture()
                Dim exitCode As Integer = processor.Run(args)
                output = capture.Output
                errorOutput = capture.ErrorOutput
                Return exitCode
            End Using

        End Function

        ''' <summary>
        ''' Runs the CLI processor with the specified arguments, discarding the console output.
        ''' <para></para>
        ''' This method intentionally has a different name than <see cref="RunCli(ByRef String, ByRef String, String())"/>:
        ''' VB.NET allows passing string literals to <see langword="ByRef"/> parameters, so an overload with the same
        ''' name would be ambiguous and the first two arguments would silently be bound to the output parameters.
        ''' </summary>
        '''
        ''' <param name="args">
        ''' The command-line arguments.
        ''' </param>
        '''
        ''' <returns>
        ''' The exit code returned by the processor.
        ''' </returns>
        Private Shared Function RunCliSilently(ParamArray args As String()) As Integer

            Dim output As String = Nothing
            Dim errorOutput As String = Nothing
            Return CommandLineProcessorTests.RunCli(output, errorOutput, args)

        End Function

#End Region

#Region " Run - General "

        ''' <summary>
        ''' Verifies that running without any action returns the argument error exit code.
        ''' </summary>
        <TestMethod>
        Public Sub Run_NoArguments_ReturnsArgumentError()

            Dim output As String = Nothing
            Dim errorOutput As String = Nothing

            Dim exitCode As Integer = CommandLineProcessorTests.RunCli(output, errorOutput)

            Assert.AreEqual(ExitArgumentError, exitCode)
            TextAssert.Contains(errorOutput, "No action specified")

        End Sub

        ''' <summary>
        ''' Verifies that a file path without any action returns the argument error exit code.
        ''' </summary>
        <TestMethod>
        Public Sub Run_OnlyFilePath_ReturnsArgumentError()

            Dim output As String = Nothing
            Dim errorOutput As String = Nothing

            Dim exitCode As Integer = CommandLineProcessorTests.RunCli(output, errorOutput, Me.tempDir.Combine("Test.lnk"))

            Assert.AreEqual(ExitArgumentError, exitCode)
            TextAssert.Contains(errorOutput, "No action specified")

        End Sub

        ''' <summary>
        ''' Verifies that parser errors are reported as argument errors.
        ''' </summary>
        <TestMethod>
        <DataRow("/Bogus", "/Read")>
        <DataRow("/Read", "/Create")>
        <DataRow("/Create", "/Hotkey:NotAKey")>
        <DataRow("/Create", "/WindowState:2")>
        <DataRow("/Create", "/IconIndex:x")>
        <DataRow("/Create", "/Target")>
        Public Sub Run_ParserError_ReturnsArgumentError(firstArgument As String, secondArgument As String)

            Dim output As String = Nothing
            Dim errorOutput As String = Nothing

            Dim exitCode As Integer = CommandLineProcessorTests.RunCli(output, errorOutput, firstArgument, secondArgument)

            Assert.AreEqual(ExitArgumentError, exitCode)
            TextAssert.Contains(errorOutput, "Argument error:")
            Assert.AreEqual(String.Empty, output)

        End Sub

        ''' <summary>
        ''' Verifies that the help text is printed and the success exit code is returned.
        ''' </summary>
        <TestMethod>
        <DataRow("/Help")>
        <DataRow("/?")>
        Public Sub Run_Help_PrintsUsageAndReturnsSuccess(helpSwitch As String)

            Dim output As String = Nothing
            Dim errorOutput As String = Nothing

            Dim exitCode As Integer = CommandLineProcessorTests.RunCli(output, errorOutput, helpSwitch)

            Assert.AreEqual(ExitSuccess, exitCode)
            Assert.AreEqual(String.Empty, errorOutput)
            TextAssert.Contains(output, "Usage:")
            TextAssert.Contains(output, "/Create")
            TextAssert.Contains(output, "/Modify")
            TextAssert.Contains(output, "/Read")
            TextAssert.Contains(output, "/WindowState:state")
            TextAssert.Contains(output, "Exit Codes:")

        End Sub

#End Region

#Region " Run - File Path Validation "

        ''' <summary>
        ''' Verifies that every action that operates on a file requires the /File switch.
        ''' </summary>
        <TestMethod>
        <DataRow("/Create", "Create")>
        <DataRow("/Modify", "Modify")>
        <DataRow("/Read", "Read")>
        Public Sub Run_ActionWithoutFile_ReturnsArgumentError(actionSwitch As String, actionName As String)

            Dim output As String = Nothing
            Dim errorOutput As String = Nothing

            Dim exitCode As Integer = CommandLineProcessorTests.RunCli(output, errorOutput, actionSwitch)

            Assert.AreEqual(ExitArgumentError, exitCode)
            TextAssert.Contains(errorOutput, $"/File is required for /{actionName}")

        End Sub

        ''' <summary>
        ''' Verifies that an empty /File value is treated as a missing file path.
        ''' </summary>
        <TestMethod>
        Public Sub Run_EmptyFileValue_ReturnsArgumentError()

            Dim output As String = Nothing
            Dim errorOutput As String = Nothing

            Dim exitCode As Integer = CommandLineProcessorTests.RunCli(output, errorOutput, "/Read", "/File:")

            Assert.AreEqual(ExitArgumentError, exitCode)
            TextAssert.Contains(errorOutput, "/File is required")

        End Sub

        ''' <summary>
        ''' Verifies that every action rejects file paths without the .lnk extension.
        ''' </summary>
        <TestMethod>
        <DataRow("/Create")>
        <DataRow("/Modify")>
        <DataRow("/Read")>
        Public Sub Run_NonLnkExtension_ReturnsArgumentError(actionSwitch As String)

            Dim output As String = Nothing
            Dim errorOutput As String = Nothing
            Dim filePath As String = Me.tempDir.Combine("NotAShortcut.txt")

            Dim exitCode As Integer = CommandLineProcessorTests.RunCli(output, errorOutput, actionSwitch, $"/File:{filePath}")

            Assert.AreEqual(ExitArgumentError, exitCode)
            TextAssert.Contains(errorOutput, ".lnk extension")
            Assert.IsFalse(File.Exists(filePath))

        End Sub

        ''' <summary>
        ''' Verifies the file path validation helper directly, including the case-insensitive extension check.
        ''' </summary>
        <TestMethod>
        <DataRow("C:\Shortcut.lnk", 0)>
        <DataRow("C:\Shortcut.LNK", 0)>
        <DataRow("Relative.Lnk", 0)>
        <DataRow("C:\Shortcut.lnk.txt", 1)>
        <DataRow("C:\Shortcut", 1)>
        <DataRow("   ", 1)>
        <DataRow("", 1)>
        Public Sub ValidateFilePath_ReturnsExpectedCode(filePath As String, expectedCode As Integer)

            Dim processor As New CommandLineProcessor()
            Dim options As New CommandLineOptions() With {.FilePath = filePath}

            Using capture As New ConsoleCapture()
                Dim result As Integer = processor.ValidateFilePath(options, "Read")
                Assert.AreEqual(expectedCode, result)
            End Using

        End Sub

        ''' <summary>
        ''' Verifies that reading a shortcut that does not exist returns the runtime error exit code.
        ''' </summary>
        <TestMethod>
        Public Sub Run_ReadMissingFile_ReturnsRuntimeError()

            Dim output As String = Nothing
            Dim errorOutput As String = Nothing
            Dim filePath As String = Me.tempDir.Combine("Missing.lnk")

            Dim exitCode As Integer = CommandLineProcessorTests.RunCli(output, errorOutput, "/Read", $"/File:{filePath}")

            Assert.AreEqual(ExitRuntimeError, exitCode)
            TextAssert.Contains(errorOutput, "File not found")

        End Sub

        ''' <summary>
        ''' Verifies that modifying a shortcut that does not exist returns the runtime error exit code,
        ''' and does not create the file.
        ''' </summary>
        <TestMethod>
        Public Sub Run_ModifyMissingFile_ReturnsRuntimeErrorAndDoesNotCreateFile()

            Dim output As String = Nothing
            Dim errorOutput As String = Nothing
            Dim filePath As String = Me.tempDir.Combine("Missing.lnk")

            Dim exitCode As Integer = CommandLineProcessorTests.RunCli(output, errorOutput, "/Modify", $"/File:{filePath}", "/Desc:Test")

            Assert.AreEqual(ExitRuntimeError, exitCode)
            TextAssert.Contains(errorOutput, "File not found")
            Assert.IsFalse(File.Exists(filePath))

        End Sub

#End Region

#Region " Run - Create (Integration) "

        ''' <summary>
        ''' Verifies that /Create writes a new shortcut with every specified property.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub Run_Create_CreatesShortcutWithAllProperties()

            Dim targetPath As String = Me.tempDir.CreateFile("Target\App.txt", "target")
            Dim iconPath As String = Me.tempDir.CreateFile("Icons\App.ico", "icon")
            Dim workDir As String = Path.GetDirectoryName(targetPath)
            Dim lnkPath As String = Me.tempDir.Combine("Created.lnk")

            Dim output As String = Nothing
            Dim errorOutput As String = Nothing

            Dim exitCode As Integer = CommandLineProcessorTests.RunCli(
                output, errorOutput,
                "/Create",
                $"/File:{lnkPath}",
                $"/Target:{targetPath}",
                "/Args:--verbose ""quoted arg""",
                "/Desc:Created by the CLI tests",
                $"/Icon:{iconPath}",
                "/IconIndex:2",
                "/Hotkey:Control+Alt+F7",
                "/WindowState:Max",
                $"/WorkDir:{workDir}")

            Assert.AreEqual(ExitSuccess, exitCode, errorOutput)
            Assert.AreEqual(String.Empty, errorOutput)
            TextAssert.Contains(output, "Shortcut created:")
            Assert.IsTrue(File.Exists(lnkPath))

            Dim shortcut As New ShortcutFileInfo(lnkPath)
            TextAssert.PathEquals(targetPath, shortcut.Target)
            Assert.AreEqual("--verbose ""quoted arg""", shortcut.TargetArguments)
            Assert.AreEqual("Created by the CLI tests", shortcut.Description)
            TextAssert.PathEquals(iconPath, shortcut.Icon)
            Assert.AreEqual(2, shortcut.IconIndex)
            Assert.AreEqual(Keys.Control Or Keys.Alt Or Keys.F7, shortcut.Hotkey)
            Assert.AreEqual(ShortcutWindowState.Maximized, shortcut.WindowState)
            TextAssert.PathEquals(workDir, shortcut.WorkingDirectory)

        End Sub

        ''' <summary>
        ''' Verifies that /Create persists the Application User Model ID.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub Run_CreateWithAppId_PersistsAppId()

            Dim targetPath As String = Me.tempDir.CreateFile("App.txt", "target")
            Dim lnkPath As String = Me.tempDir.Combine("WithAppId.lnk")
            Const appId As String = "ElektroStudios.EasyLinkFileViewer.Tests"

            Dim output As String = Nothing
            Dim errorOutput As String = Nothing

            Dim exitCode As Integer = CommandLineProcessorTests.RunCli(
                output, errorOutput, "/Create", $"/File:{lnkPath}", $"/Target:{targetPath}", $"/AppId:{appId}")

            Assert.AreEqual(ExitSuccess, exitCode, errorOutput)
            Assert.AreEqual(appId, ShortcutAppIdHelper.GetAppUserModelId(lnkPath))

        End Sub

        ''' <summary>
        ''' Verifies that /Create refuses to overwrite an existing file without /Force,
        ''' and leaves the existing file untouched.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub Run_CreateExistingWithoutForce_ReturnsRuntimeErrorAndKeepsFile()

            Dim targetPath As String = Me.tempDir.CreateFile("App.txt", "target")
            Dim lnkPath As String = Me.tempDir.Combine("Existing.lnk")
            Assert.AreEqual(ExitSuccess, CommandLineProcessorTests.RunCliSilently("/Create", $"/File:{lnkPath}", $"/Target:{targetPath}", "/Desc:Original"))

            Dim output As String = Nothing
            Dim errorOutput As String = Nothing

            Dim exitCode As Integer = CommandLineProcessorTests.RunCli(
                output, errorOutput, "/Create", $"/File:{lnkPath}", $"/Target:{targetPath}", "/Desc:Replaced")

            Assert.AreEqual(ExitRuntimeError, exitCode)
            TextAssert.Contains(errorOutput, "File already exists")
            TextAssert.Contains(errorOutput, "/Force")
            Assert.AreEqual("Original", New ShortcutFileInfo(lnkPath).Description)

        End Sub

        ''' <summary>
        ''' Verifies that /Create /Force creates the shortcut from scratch, without carrying over
        ''' properties of the overwritten shortcut that were not specified on the command line
        ''' (regression test).
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub Run_CreateExistingWithForce_OverwritesFromScratch()

            Dim oldTargetPath As String = Me.tempDir.CreateFile("Old.txt", "old")
            Dim newTargetPath As String = Me.tempDir.CreateFile("New.txt", "new")
            Dim lnkPath As String = Me.tempDir.Combine("Overwritten.lnk")

            Assert.AreEqual(ExitSuccess, CommandLineProcessorTests.RunCliSilently(
                "/Create", $"/File:{lnkPath}", $"/Target:{oldTargetPath}",
                "/Desc:Old description", "/Args:--old", "/Hotkey:Control+Alt+O", "/WindowState:Min"))

            Dim output As String = Nothing
            Dim errorOutput As String = Nothing

            Dim exitCode As Integer = CommandLineProcessorTests.RunCli(
                output, errorOutput, "/Create", "/Force", $"/File:{lnkPath}", $"/Target:{newTargetPath}")

            Assert.AreEqual(ExitSuccess, exitCode, errorOutput)

            Dim shortcut As New ShortcutFileInfo(lnkPath)
            TextAssert.PathEquals(newTargetPath, shortcut.Target)
            Assert.AreEqual(String.Empty, shortcut.Description)
            Assert.AreEqual(String.Empty, shortcut.TargetArguments)
            Assert.AreEqual(Keys.None, shortcut.Hotkey)
            Assert.AreEqual(ShortcutWindowState.Normal, shortcut.WindowState)

        End Sub

        ''' <summary>
        ''' Verifies that /Create /Force can overwrite a read-only shortcut file.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub Run_CreateReadOnlyExistingWithForce_Overwrites()

            Dim targetPath As String = Me.tempDir.CreateFile("App.txt", "target")
            Dim lnkPath As String = Me.tempDir.Combine("ReadOnly.lnk")
            Assert.AreEqual(ExitSuccess, CommandLineProcessorTests.RunCliSilently("/Create", $"/File:{lnkPath}", $"/Target:{targetPath}", "/Desc:Original"))
            File.SetAttributes(lnkPath, File.GetAttributes(lnkPath) Or FileAttributes.ReadOnly)

            Dim output As String = Nothing
            Dim errorOutput As String = Nothing

            Dim exitCode As Integer = CommandLineProcessorTests.RunCli(
                output, errorOutput, "/Create", "/Force", $"/File:{lnkPath}", $"/Target:{targetPath}", "/Desc:Replaced")

            Assert.AreEqual(ExitSuccess, exitCode, errorOutput)
            Assert.AreEqual("Replaced", New ShortcutFileInfo(lnkPath).Description)

        End Sub

        ''' <summary>
        ''' Verifies that /Create without /WindowState stores a normal window state,
        ''' instead of the native SW_HIDE value (0) that would launch the target hidden (regression test).
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub Run_CreateWithoutWindowState_StoresNormal()

            Dim targetPath As String = Me.tempDir.CreateFile("App.txt", "target")
            Dim lnkPath As String = Me.tempDir.Combine("DefaultWindowState.lnk")

            Assert.AreEqual(ExitSuccess, CommandLineProcessorTests.RunCliSilently("/Create", $"/File:{lnkPath}", $"/Target:{targetPath}"))

            Assert.AreEqual(ShortcutWindowState.Normal, New ShortcutFileInfo(lnkPath).WindowState)

        End Sub

#End Region

#Region " Run - Modify (Integration) "

        ''' <summary>
        ''' Verifies that /Modify changes only the specified properties and preserves the rest.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub Run_Modify_ChangesOnlySpecifiedProperties()

            Dim targetPath As String = Me.tempDir.CreateFile("App.txt", "target")
            Dim iconPath As String = Me.tempDir.CreateFile("App.ico", "icon")
            Dim lnkPath As String = Me.tempDir.Combine("Modified.lnk")

            Assert.AreEqual(ExitSuccess, CommandLineProcessorTests.RunCliSilently(
                "/Create", $"/File:{lnkPath}", $"/Target:{targetPath}",
                "/Args:--keep", "/Desc:Before", "/Hotkey:Control+Shift+K", "/WindowState:Min",
                $"/Icon:{iconPath}", "/IconIndex:4"))

            Dim output As String = Nothing
            Dim errorOutput As String = Nothing

            Dim exitCode As Integer = CommandLineProcessorTests.RunCli(
                output, errorOutput, "/Modify", $"/File:{lnkPath}", "/Desc:After", "/WindowState:Normal")

            Assert.AreEqual(ExitSuccess, exitCode, errorOutput)
            TextAssert.Contains(output, "Shortcut modified:")

            Dim shortcut As New ShortcutFileInfo(lnkPath)
            Assert.AreEqual("After", shortcut.Description)
            Assert.AreEqual(ShortcutWindowState.Normal, shortcut.WindowState)

            ' Untouched properties.
            TextAssert.PathEquals(targetPath, shortcut.Target)
            Assert.AreEqual("--keep", shortcut.TargetArguments)
            Assert.AreEqual(Keys.Control Or Keys.Shift Or Keys.K, shortcut.Hotkey)
            TextAssert.PathEquals(iconPath, shortcut.Icon)
            Assert.AreEqual(4, shortcut.IconIndex)

        End Sub

        ''' <summary>
        ''' Verifies that /Modify with an empty value clears that property.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub Run_ModifyWithEmptyValue_ClearsProperty()

            Dim targetPath As String = Me.tempDir.CreateFile("App.txt", "target")
            Dim lnkPath As String = Me.tempDir.Combine("Cleared.lnk")
            Assert.AreEqual(ExitSuccess, CommandLineProcessorTests.RunCliSilently(
                "/Create", $"/File:{lnkPath}", $"/Target:{targetPath}", "/Desc:To be cleared", "/Args:--to-be-cleared"))

            Assert.AreEqual(ExitSuccess, CommandLineProcessorTests.RunCliSilently("/Modify", $"/File:{lnkPath}", "/Desc:", "/Args:"))

            Dim shortcut As New ShortcutFileInfo(lnkPath)
            Assert.AreEqual(String.Empty, shortcut.Description)
            Assert.AreEqual(String.Empty, shortcut.TargetArguments)

        End Sub

        ''' <summary>
        ''' Verifies that /Modify with /Hotkey:None removes the hotkey.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub Run_ModifyHotkeyNone_RemovesHotkey()

            Dim targetPath As String = Me.tempDir.CreateFile("App.txt", "target")
            Dim lnkPath As String = Me.tempDir.Combine("Hotkey.lnk")
            Assert.AreEqual(ExitSuccess, CommandLineProcessorTests.RunCliSilently(
                "/Create", $"/File:{lnkPath}", $"/Target:{targetPath}", "/Hotkey:Control+Alt+H"))

            Assert.AreEqual(ExitSuccess, CommandLineProcessorTests.RunCliSilently("/Modify", $"/File:{lnkPath}", "/Hotkey:None"))

            Assert.AreEqual(Keys.None, New ShortcutFileInfo(lnkPath).Hotkey)

        End Sub

#End Region

#Region " Run - Read (Integration) "

        ''' <summary>
        ''' Verifies that /Read prints every property of the shortcut.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub Run_Read_PrintsShortcutProperties()

            Dim targetPath As String = Me.tempDir.CreateFile("App.txt", "target")
            Dim lnkPath As String = Me.tempDir.Combine("Readable.lnk")
            Assert.AreEqual(ExitSuccess, CommandLineProcessorTests.RunCliSilently(
                "/Create", $"/File:{lnkPath}", $"/Target:{targetPath}",
                "/Desc:Readable description", "/Args:--read-me", "/Hotkey:Control+Alt+F1", "/WindowState:Max"))

            Dim output As String = Nothing
            Dim errorOutput As String = Nothing

            Dim exitCode As Integer = CommandLineProcessorTests.RunCli(output, errorOutput, "/Read", $"/File:{lnkPath}")

            Assert.AreEqual(ExitSuccess, exitCode, errorOutput)
            Assert.AreEqual(String.Empty, errorOutput)
            TextAssert.Contains(output, "Shortcut Properties")
            TextAssert.Contains(output, "  File Name        : Readable.lnk")
            TextAssert.Contains(output, "  Description      : Readable description")
            TextAssert.Contains(output, "  Target Arguments : --read-me")
            TextAssert.Contains(output, "  Hotkey           : F1 + Control + Alt")
            TextAssert.Contains(output, "  Window State     : Maximized")
            TextAssert.Contains(output, "App.txt")

        End Sub

        ''' <summary>
        ''' Verifies that /Read accepts the shortcut path as a positional argument.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub Run_ReadWithPositionalPath_Succeeds()

            Dim targetPath As String = Me.tempDir.CreateFile("App.txt", "target")
            Dim lnkPath As String = Me.tempDir.Combine("Positional.lnk")
            Assert.AreEqual(ExitSuccess, CommandLineProcessorTests.RunCliSilently("/Create", $"/File:{lnkPath}", $"/Target:{targetPath}"))

            Dim output As String = Nothing
            Dim errorOutput As String = Nothing

            Dim exitCode As Integer = CommandLineProcessorTests.RunCli(output, errorOutput, lnkPath, "/Read")

            Assert.AreEqual(ExitSuccess, exitCode, errorOutput)
            TextAssert.Contains(output, "  Hotkey           : None")

        End Sub

#End Region

#Region " FormatHotkey "

        ''' <summary>
        ''' Verifies the human-readable hotkey format used by the /Read action.
        ''' </summary>
        <TestMethod>
        <DataRow(Keys.None, "None")>
        <DataRow(Keys.F1, "F1")>
        <DataRow(Keys.Control Or Keys.Alt Or Keys.F1, "F1 + Control + Alt")>
        <DataRow(Keys.Shift Or Keys.A, "A + Shift")>
        Public Sub FormatHotkey_ReturnsReadableText(hotkey As Keys, expected As String)

            Dim result As String = CommandLineProcessor.FormatHotkey(hotkey)

            Assert.AreEqual(expected, result)

        End Sub

#End Region

    End Class

End Namespace

#End Region
