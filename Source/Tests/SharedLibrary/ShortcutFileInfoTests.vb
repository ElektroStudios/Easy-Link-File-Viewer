#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Imports DevCase.Core.IO

Imports EasyLink.Tests.Helpers

Imports Microsoft.VisualStudio.TestTools.UnitTesting

#End Region

#Region " ShortcutFileInfoTests "

Namespace EasyLink.Tests.SharedLibrary

    ''' <summary>
    ''' Unit and integration tests for the <see cref="ShortcutFileInfo"/> class.
    ''' <para></para>
    ''' The integration tests create and read real shortcut files inside a temporary directory
    ''' through the Windows Shell COM objects, so they run on a single-threaded apartment (STA) thread.
    ''' </summary>
    <STATestClass>
    Public NotInheritable Class ShortcutFileInfoTests

#Region " Private Fields "

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
        ''' Creates a real shortcut file in the temporary directory that points to a new text file.
        ''' </summary>
        '''
        ''' <param name="lnkName">
        ''' The file name of the shortcut to create.
        ''' </param>
        '''
        ''' <param name="targetPath">
        ''' Receives the full path of the created target file.
        ''' </param>
        '''
        ''' <returns>
        ''' The full path of the created shortcut file.
        ''' </returns>
        Private Function CreateShortcut(lnkName As String, ByRef targetPath As String) As String

            targetPath = Me.tempDir.CreateFile($"Targets\{Path.GetFileNameWithoutExtension(lnkName)}.txt", "target")
            Dim lnkPath As String = Me.tempDir.Combine(lnkName)

            Dim shortcut As New ShortcutFileInfo(lnkPath) With {.ViewMode = True}
            shortcut.Target = targetPath
            shortcut.Description = "Test shortcut"
            shortcut.Create()

            Return lnkPath

        End Function

        ''' <summary>
        ''' Creates a real shortcut file in the temporary directory that points to a new text file.
        ''' </summary>
        '''
        ''' <param name="lnkName">
        ''' The file name of the shortcut to create.
        ''' </param>
        '''
        ''' <returns>
        ''' The full path of the created shortcut file.
        ''' </returns>
        Private Function CreateShortcut(lnkName As String) As String

            Dim targetPath As String = Nothing
            Return Me.CreateShortcut(lnkName, targetPath)

        End Function

        ''' <summary>
        ''' Builds the minimal bytes of a shell link binary header (MS-SHLLINK) with the specified link flags,
        ''' optionally followed by a LinkTargetIDList block.
        ''' </summary>
        '''
        ''' <param name="linkFlags">
        ''' The value of the LinkFlags header field.
        ''' </param>
        '''
        ''' <param name="idListSize">
        ''' The value of the IDListSize field that follows the header.
        ''' </param>
        '''
        ''' <param name="idListBytes">
        ''' The raw IDList bytes appended after the IDListSize field.
        ''' </param>
        '''
        ''' <returns>
        ''' The resulting bytes.
        ''' </returns>
        Private Shared Function BuildLnkBinary(linkFlags As UInteger, idListSize As UShort, idListBytes As Byte()) As Byte()

            Const HeaderSize As Integer = &H4C
            Const LinkFlagsOffset As Integer = &H14

            Dim bytes As Byte() = New Byte(HeaderSize + 2 + idListBytes.Length - 1) {}
            Array.Copy(BitConverter.GetBytes(CUInt(HeaderSize)), 0, bytes, 0, 4)
            Array.Copy(BitConverter.GetBytes(linkFlags), 0, bytes, LinkFlagsOffset, 4)
            Array.Copy(BitConverter.GetBytes(idListSize), 0, bytes, HeaderSize, 2)
            Array.Copy(idListBytes, 0, bytes, HeaderSize + 2, idListBytes.Length)
            Return bytes

        End Function

#End Region

#Region " File Info Properties (Pure) "

        ''' <summary>
        ''' Verifies the file-related properties of an instance whose file does not exist.
        ''' </summary>
        <TestMethod>
        Public Sub Constructor_NonExistingFile_ExposesFileInfoProperties()

            Dim lnkPath As String = Me.tempDir.Combine("Missing.lnk")

            Dim shortcut As New ShortcutFileInfo(lnkPath)

            Assert.AreEqual("Missing.lnk", shortcut.Name)
            Assert.AreEqual(lnkPath, shortcut.FullName)
            Assert.AreEqual(".lnk", shortcut.Extension)
            Assert.AreEqual(Me.tempDir.FullPath, shortcut.DirectoryName)
            Assert.AreEqual(Me.tempDir.FullPath, shortcut.Directory.FullName)
            Assert.IsFalse(shortcut.Exists)
            Assert.AreEqual(lnkPath, shortcut.ToString())

        End Sub

        ''' <summary>
        ''' Verifies that a relative path is resolved to a full path, while <see cref="ShortcutFileInfo.ToString()"/>
        ''' keeps returning the original path.
        ''' </summary>
        <TestMethod>
        Public Sub Constructor_RelativePath_ResolvesFullPathAndKeepsOriginalPath()

            Const relativePath As String = "EasyLink.Tests.Relative.Missing.lnk"

            Dim shortcut As New ShortcutFileInfo(relativePath)

            Assert.AreEqual(Path.GetFullPath(relativePath), shortcut.FullName)
            Assert.AreEqual(relativePath, shortcut.ToString())

        End Sub

        ''' <summary>
        ''' Verifies the default values of an instance whose file does not exist.
        ''' </summary>
        <TestMethod>
        Public Sub Constructor_NonExistingFile_HasDefaultShortcutValues()

            Dim shortcut As New ShortcutFileInfo(Me.tempDir.Combine("Missing.lnk"))

            Assert.AreEqual(Keys.None, shortcut.Hotkey)
            Assert.AreEqual(0, shortcut.IconIndex)
            Assert.AreEqual(ShortcutWindowState.Normal, shortcut.WindowState)
            Assert.AreEqual(ShortcutTargetExecutionMode.Standard, shortcut.TargetExecutionMode)
            Assert.IsFalse(shortcut.ViewMode)
            Assert.IsFalse(shortcut.IsWindowsInstallerShortcut)
            Assert.AreEqual(IntPtr.Zero, shortcut.TargetPidl)
            Assert.AreEqual(String.Empty, shortcut.TargetDisplayName)

        End Sub

        ''' <summary>
        ''' Verifies that the <see cref="FileInfo"/> constructor overload is equivalent to the path overload.
        ''' </summary>
        <TestMethod>
        Public Sub Constructor_FileInfoOverload_UsesSameFullPath()

            Dim lnkPath As String = Me.tempDir.Combine("Missing.lnk")

            Dim shortcut As New ShortcutFileInfo(New FileInfo(lnkPath))

            Assert.AreEqual(lnkPath, shortcut.FullName)

        End Sub

#End Region

#Region " ViewMode and Property Setters (Pure) "

        ''' <summary>
        ''' Verifies that, outside view mode, assigning a shortcut property of a non-existing file
        ''' throws a <see cref="FileNotFoundException"/> because changes are written immediately.
        ''' </summary>
        <TestMethod>
        Public Sub PropertySetter_NotViewModeAndFileMissing_ThrowsFileNotFoundException()

            Dim shortcut As New ShortcutFileInfo(Me.tempDir.Combine("Missing.lnk"))

            Assert.ThrowsExactly(Of FileNotFoundException)(Sub() shortcut.Description = "Some description")

        End Sub

        ''' <summary>
        ''' Verifies that, in view mode, property assignments are held in memory and nothing is written to disk.
        ''' </summary>
        <TestMethod>
        Public Sub PropertySetter_ViewMode_KeepsValuesInMemoryWithoutWriting()

            Dim lnkPath As String = Me.tempDir.Combine("InMemory.lnk")
            Dim shortcut As New ShortcutFileInfo(lnkPath) With {.ViewMode = True}

            shortcut.Target = "C:\Windows\notepad.exe"
            shortcut.TargetArguments = "--arg"
            shortcut.Description = "Description"
            shortcut.Icon = "C:\Windows\System32\shell32.dll"
            shortcut.IconIndex = 7
            shortcut.Hotkey = Keys.Control Or Keys.Alt Or Keys.Z
            shortcut.WindowState = ShortcutWindowState.Minimized
            shortcut.WorkingDirectory = "C:\Windows"
            shortcut.TargetExecutionMode = ShortcutTargetExecutionMode.RelativePathViaExplorer

            Assert.AreEqual("C:\Windows\notepad.exe", shortcut.Target)
            Assert.AreEqual("--arg", shortcut.TargetArguments)
            Assert.AreEqual("Description", shortcut.Description)
            Assert.AreEqual("C:\Windows\System32\shell32.dll", shortcut.Icon)
            Assert.AreEqual(7, shortcut.IconIndex)
            Assert.AreEqual(Keys.Control Or Keys.Alt Or Keys.Z, shortcut.Hotkey)
            Assert.AreEqual(ShortcutWindowState.Minimized, shortcut.WindowState)
            Assert.AreEqual("C:\Windows", shortcut.WorkingDirectory)
            Assert.AreEqual(ShortcutTargetExecutionMode.RelativePathViaExplorer, shortcut.TargetExecutionMode)
            Assert.IsFalse(File.Exists(lnkPath))

        End Sub

        ''' <summary>
        ''' Verifies that the string properties are truncated to the maximum length supported by the Shell link format.
        ''' </summary>
        <TestMethod>
        Public Sub PropertySetter_ViewModeTooLongValues_AreTruncated()

            Dim shortcut As New ShortcutFileInfo(Me.tempDir.Combine("Truncated.lnk")) With {.ViewMode = True}

            shortcut.Description = New String("d"c, 500)
            shortcut.Icon = New String("i"c, 500)
            shortcut.Target = New String("t"c, 500)
            shortcut.WorkingDirectory = New String("w"c, 500)
            shortcut.TargetArguments = New String("a"c, 5000)

            Assert.AreEqual(259, shortcut.Description.Length)
            Assert.AreEqual(259, shortcut.Icon.Length)
            Assert.AreEqual(259, shortcut.Target.Length)
            Assert.AreEqual(259, shortcut.WorkingDirectory.Length)
            Assert.AreEqual(1023, shortcut.TargetArguments.Length)

        End Sub

        ''' <summary>
        ''' Verifies that values at exactly the maximum length are kept intact.
        ''' </summary>
        <TestMethod>
        Public Sub PropertySetter_ViewModeMaxLengthValues_AreNotTruncated()

            Dim shortcut As New ShortcutFileInfo(Me.tempDir.Combine("MaxLength.lnk")) With {.ViewMode = True}
            Dim description As String = New String("d"c, 259)
            Dim arguments As String = New String("a"c, 1023)

            shortcut.Description = description
            shortcut.TargetArguments = arguments

            Assert.AreEqual(description, shortcut.Description)
            Assert.AreEqual(arguments, shortcut.TargetArguments)

        End Sub

        ''' <summary>
        ''' Verifies that the AppId setter treats <see langword="Nothing"/> as an empty string.
        ''' </summary>
        <TestMethod>
        Public Sub AppId_ViewModeSetNothing_BecomesEmptyString()

            Dim shortcut As New ShortcutFileInfo(Me.tempDir.Combine("AppId.lnk")) With {.ViewMode = True}

            shortcut.AppId = "Company.Product"
            Assert.AreEqual("Company.Product", shortcut.AppId)

            shortcut.AppId = Nothing
            Assert.AreEqual(String.Empty, shortcut.AppId)

        End Sub

#End Region

#Region " Hotkey Conversion (Pure) "

        ''' <summary>
        ''' Verifies the conversion from <see cref="Keys"/> to the native IShellLink hotkey word
        ''' (low byte: virtual key code; high byte: HOTKEYF_* modifiers).
        ''' </summary>
        <TestMethod>
        <DataRow(Keys.None, 0)>
        <DataRow(Keys.F1, &H70)>
        <DataRow(Keys.Shift Or Keys.A, &H141)>
        <DataRow(Keys.Control Or Keys.Alt Or Keys.F1, &H670)>
        <DataRow(Keys.Control Or Keys.Shift Or Keys.Alt Or Keys.Z, &H75A)>
        Public Sub KeysToHotkey_ReturnsNativeHotkeyValue(key As Keys, expected As Integer)

            Dim shortcut As New ShortcutFileInfo(Me.tempDir.Combine("Hotkey.lnk"))

            Dim result As UShort = shortcut.KeysToHotkey(key)

            Assert.AreEqual(CUShort(expected), result)

        End Sub

        ''' <summary>
        ''' Verifies the conversion from a native modifier and key code to a <see cref="Keys"/> value.
        ''' </summary>
        <TestMethod>
        <DataRow(ShortcutHotkeyModifier.None, Keys.F1, Keys.F1)>
        <DataRow(ShortcutHotkeyModifier.Shift, Keys.A, Keys.Shift Or Keys.A)>
        <DataRow(ShortcutHotkeyModifier.Control Or ShortcutHotkeyModifier.Alt, Keys.F1, Keys.Control Or Keys.Alt Or Keys.F1)>
        <DataRow(ShortcutHotkeyModifier.Shift Or ShortcutHotkeyModifier.Control Or ShortcutHotkeyModifier.Alt, Keys.Z, Keys.Shift Or Keys.Control Or Keys.Alt Or Keys.Z)>
        Public Sub HotkeyToKeys_ReturnsKeysValue(modifier As ShortcutHotkeyModifier, keyCode As Keys, expected As Keys)

            Dim shortcut As New ShortcutFileInfo(Me.tempDir.Combine("Hotkey.lnk"))

            Dim result As Keys = shortcut.HotkeyToKeys(modifier, keyCode)

            Assert.AreEqual(expected, result)

        End Sub

        ''' <summary>
        ''' Verifies that converting a <see cref="Keys"/> value to the native hotkey word and back is lossless.
        ''' </summary>
        <TestMethod>
        <DataRow(Keys.F12)>
        <DataRow(Keys.Control Or Keys.D1)>
        <DataRow(Keys.Alt Or Keys.Shift Or Keys.Home)>
        <DataRow(Keys.Control Or Keys.Alt Or Keys.Shift Or Keys.NumPad5)>
        Public Sub HotkeyConversion_RoundTrip_IsLossless(key As Keys)

            Dim shortcut As New ShortcutFileInfo(Me.tempDir.Combine("Hotkey.lnk"))

            Dim nativeValue As UShort = shortcut.KeysToHotkey(key)
            Dim bytes As Byte() = BitConverter.GetBytes(nativeValue)
            Dim result As Keys = shortcut.HotkeyToKeys(CType(bytes(1), ShortcutHotkeyModifier), CType(bytes(0), Keys))

            Assert.AreEqual(key, result)

        End Sub

#End Region

#Region " IsShellNamespaceTarget (Pure) "

        ''' <summary>
        ''' Verifies the detection of virtual Shell namespace targets (CLSID paths, UWP AppUserModelIDs and "shell:" URIs).
        ''' </summary>
        <TestMethod>
        <DataRow("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", True)>
        <DataRow("{20D04FE0-3AEA-1069-A2D8-08002B30309D}", True)>
        <DataRow("::{82E749ED-B971-4550-BAF7-06AA2BF7E836}\AAA_SystemSettings_Connections_Hotspot2_SignUp_Toggle", True)>
        <DataRow("windows.immersivecontrolpanel_cw5n1h2txyewy!microsoft.windows.immersivecontrolpanel", True)>
        <DataRow("Microsoft.WindowsCalculator_8wekyb3d8bbwe!App", True)>
        <DataRow("shell:AppsFolder", True)>
        <DataRow("SHELL:Downloads", True)>
        <DataRow("C:\Windows\notepad.exe", False)>
        <DataRow("\\Server\Share\File.txt", False)>
        <DataRow("{not-a-guid}", False)>
        <DataRow("::{20D04FE0-3AEA-1069-A2D8}", False)>
        <DataRow("Microsoft.WindowsCalculator_8wekyb3d8bbwe", False)>
        <DataRow("", False)>
        Public Sub IsShellNamespaceTarget_ReturnsExpectedResult(target As String, expected As Boolean)

            Dim result As Boolean = ShortcutFileInfo.IsShellNamespaceTarget(target)

            Assert.AreEqual(expected, result)

        End Sub

        ''' <summary>
        ''' Verifies that a <see langword="Nothing"/> target is not a Shell namespace target.
        ''' </summary>
        <TestMethod>
        Public Sub IsShellNamespaceTarget_Nothing_ReturnsFalse()

            Assert.IsFalse(ShortcutFileInfo.IsShellNamespaceTarget(Nothing))

        End Sub

#End Region

#Region " ReadIDListFromLnkBinary (Pure) "

        ''' <summary>
        ''' Verifies that a file smaller than the shell link header is rejected.
        ''' </summary>
        <TestMethod>
        Public Sub ReadIDListFromLnkBinary_FileTooSmall_ThrowsInvalidDataException()

            Dim filePath As String = Me.tempDir.CreateBinaryFile("Small.lnk", New Byte(9) {})

            Assert.ThrowsExactly(Of InvalidDataException)(Sub() ShortcutFileInfo.ReadIDListFromLnkBinary(filePath))

        End Sub

        ''' <summary>
        ''' Verifies that a file with an invalid header size is rejected.
        ''' </summary>
        <TestMethod>
        Public Sub ReadIDListFromLnkBinary_InvalidHeaderSize_ThrowsInvalidDataException()

            Dim filePath As String = Me.tempDir.CreateBinaryFile("BadHeader.lnk", New Byte(&H60) {})

            Dim ex As InvalidDataException = Assert.ThrowsExactly(Of InvalidDataException)(
                Sub() ShortcutFileInfo.ReadIDListFromLnkBinary(filePath))

            TextAssert.Contains(ex.Message, "header size")

        End Sub

        ''' <summary>
        ''' Verifies that <see cref="IntPtr.Zero"/> is returned when the HasLinkTargetIDList flag is not set.
        ''' </summary>
        <TestMethod>
        Public Sub ReadIDListFromLnkBinary_NoIDListFlag_ReturnsZero()

            Dim bytes As Byte() = ShortcutFileInfoTests.BuildLnkBinary(0UI, 0US, New Byte() {})
            Dim filePath As String = Me.tempDir.CreateBinaryFile("NoIdList.lnk", bytes)

            Dim result As IntPtr = ShortcutFileInfo.ReadIDListFromLnkBinary(filePath)

            Assert.AreEqual(IntPtr.Zero, result)

        End Sub

        ''' <summary>
        ''' Verifies that a zero IDListSize is rejected when the HasLinkTargetIDList flag is set.
        ''' </summary>
        <TestMethod>
        Public Sub ReadIDListFromLnkBinary_ZeroIDListSize_ThrowsInvalidDataException()

            Dim bytes As Byte() = ShortcutFileInfoTests.BuildLnkBinary(1UI, 0US, New Byte() {0, 0})
            Dim filePath As String = Me.tempDir.CreateBinaryFile("ZeroSize.lnk", bytes)

            Assert.ThrowsExactly(Of InvalidDataException)(Sub() ShortcutFileInfo.ReadIDListFromLnkBinary(filePath))

        End Sub

        ''' <summary>
        ''' Verifies that an IDListSize that exceeds the file length is rejected.
        ''' </summary>
        <TestMethod>
        Public Sub ReadIDListFromLnkBinary_IDListSizeBeyondFileEnd_ThrowsInvalidDataException()

            Dim bytes As Byte() = ShortcutFileInfoTests.BuildLnkBinary(1UI, 500US, New Byte() {0, 0})
            Dim filePath As String = Me.tempDir.CreateBinaryFile("Overflow.lnk", bytes)

            Assert.ThrowsExactly(Of InvalidDataException)(Sub() ShortcutFileInfo.ReadIDListFromLnkBinary(filePath))

        End Sub

        ''' <summary>
        ''' Verifies that the raw IDList bytes are copied into a newly allocated PIDL.
        ''' </summary>
        <TestMethod>
        Public Sub ReadIDListFromLnkBinary_ValidIDList_ReturnsPidlWithSameBytes()

            ' One ItemID (cb = 4: 2 bytes size + 2 bytes data) followed by the 2-byte terminator.
            Dim idList As Byte() = {4, 0, &HAA, &HBB, 0, 0}
            Dim bytes As Byte() = ShortcutFileInfoTests.BuildLnkBinary(1UI, CUShort(idList.Length), idList)
            Dim filePath As String = Me.tempDir.CreateBinaryFile("Valid.lnk", bytes)

            Dim pidl As IntPtr = ShortcutFileInfo.ReadIDListFromLnkBinary(filePath)
            Try
                Assert.AreNotEqual(IntPtr.Zero, pidl)
                Dim copied As Byte() = New Byte(idList.Length - 1) {}
                Marshal.Copy(pidl, copied, 0, copied.Length)
                CollectionAssert.AreEqual(idList, copied)

            Finally
                Marshal.FreeCoTaskMem(pidl)
            End Try

        End Sub

#End Region

#Region " CloneTargetPidl (Pure) "

        ''' <summary>
        ''' Verifies that cloning a missing PIDL returns <see cref="IntPtr.Zero"/>.
        ''' </summary>
        <TestMethod>
        Public Sub CloneTargetPidl_NoPidl_ReturnsZero()

            Dim shortcut As New ShortcutFileInfo(Me.tempDir.Combine("NoPidl.lnk"))

            Assert.AreEqual(IntPtr.Zero, shortcut.CloneTargetPidl())

        End Sub

        ''' <summary>
        ''' Verifies that cloning a PIDL produces an independent copy of the whole ItemID list, terminator included.
        ''' </summary>
        <TestMethod>
        Public Sub CloneTargetPidl_WithPidl_ReturnsIndependentCopy()

            ' Two ItemIDs (cb = 4 and cb = 3) followed by the 2-byte terminator.
            Dim idList As Byte() = {4, 0, &H11, &H22, 3, 0, &H33, 0, 0}
            Dim pidl As IntPtr = Marshal.AllocCoTaskMem(idList.Length)
            Marshal.Copy(idList, 0, pidl, idList.Length)

            Dim shortcut As New ShortcutFileInfo(Me.tempDir.Combine("Pidl.lnk"))
            shortcut.TargetPidl = pidl ' The instance takes ownership of the PIDL.

            Dim clone As IntPtr = shortcut.CloneTargetPidl()
            Try
                Assert.AreNotEqual(IntPtr.Zero, clone)
                Assert.AreNotEqual(pidl, clone)

                Dim cloned As Byte() = New Byte(idList.Length - 1) {}
                Marshal.Copy(clone, cloned, 0, cloned.Length)
                CollectionAssert.AreEqual(idList, cloned)

            Finally
                Marshal.FreeCoTaskMem(clone)
                shortcut.TargetPidl = IntPtr.Zero ' Frees the original PIDL.
            End Try

        End Sub

#End Region

#Region " Equality Operators "

        ''' <summary>
        ''' Verifies that the equality operators handle <see langword="Nothing"/> operands
        ''' instead of throwing a <see cref="NullReferenceException"/> (regression test).
        ''' </summary>
        <TestMethod>
        Public Sub EqualityOperators_NothingOperands_DoNotThrow()

            Dim shortcut As New ShortcutFileInfo(Me.tempDir.Combine("Missing.lnk"))
            Dim nullShortcut As ShortcutFileInfo = Nothing

            Assert.IsTrue(nullShortcut = Nothing)
            Assert.IsFalse(nullShortcut <> Nothing)
            Assert.IsFalse(shortcut = nullShortcut)
            Assert.IsFalse(nullShortcut = shortcut)
            Assert.IsTrue(shortcut <> nullShortcut)
            Assert.IsFalse(shortcut.Equals(nullShortcut))

        End Sub

        ''' <summary>
        ''' Verifies that an instance is equal to itself.
        ''' </summary>
        <TestMethod>
        Public Sub EqualityOperators_SameInstance_AreEqual()

            Dim shortcut As New ShortcutFileInfo(Me.tempDir.Combine("Missing.lnk"))
            Dim sameReference As ShortcutFileInfo = shortcut

            Assert.IsTrue(shortcut = sameReference)
            Assert.IsFalse(shortcut <> sameReference)
            Assert.IsTrue(shortcut.Equals(sameReference))

        End Sub

        ''' <summary>
        ''' Verifies that two instances that read the same shortcut file are equal and have the same hash code.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub EqualityOperators_TwoInstancesOfSameFile_AreEqual()

            Dim lnkPath As String = Me.CreateShortcut("Equal.lnk")

            Dim first As New ShortcutFileInfo(lnkPath)
            Dim second As New ShortcutFileInfo(lnkPath)

            Assert.IsTrue(first = second)
            Assert.IsFalse(first <> second)
            Assert.AreEqual(first.GetHashCode(), second.GetHashCode())

        End Sub

        ''' <summary>
        ''' Verifies that two shortcuts with different properties are not equal.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub EqualityOperators_DifferentShortcuts_AreNotEqual()

            Dim first As New ShortcutFileInfo(Me.CreateShortcut("First.lnk"))
            Dim second As New ShortcutFileInfo(Me.CreateShortcut("Second.lnk"))

            Assert.IsFalse(first = second)
            Assert.IsTrue(first <> second)

        End Sub

#End Region

#Region " Create / Read Round-Trip (Integration) "

        ''' <summary>
        ''' Verifies that every shortcut property written by <see cref="ShortcutFileInfo.Create()"/>
        ''' is read back by a new instance.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub Create_AllProperties_RoundTrip()

            Dim targetPath As String = Me.tempDir.CreateFile("Target\Program.txt", "target")
            Dim iconPath As String = Me.tempDir.CreateFile("Icons\Program.ico", "icon")
            Dim workDir As String = Path.GetDirectoryName(targetPath)
            Dim lnkPath As String = Me.tempDir.Combine("RoundTrip.lnk")

            Dim written As New ShortcutFileInfo(lnkPath) With {.ViewMode = True}
            written.Target = targetPath
            written.TargetArguments = "--flag ""value with spaces"""
            written.Description = "Round-trip description"
            written.Icon = iconPath
            written.IconIndex = 3
            written.Hotkey = Keys.Control Or Keys.Shift Or Keys.F9
            written.WindowState = ShortcutWindowState.Minimized
            written.WorkingDirectory = workDir
            written.Create()

            Assert.IsTrue(File.Exists(lnkPath))
            Assert.IsTrue(written.Exists)
            Assert.IsTrue(written.Length > 0L)

            Dim read As New ShortcutFileInfo(lnkPath)
            TextAssert.PathEquals(targetPath, read.Target)
            Assert.AreEqual("--flag ""value with spaces""", read.TargetArguments)
            Assert.AreEqual("Round-trip description", read.Description)
            TextAssert.PathEquals(iconPath, read.Icon)
            Assert.AreEqual(3, read.IconIndex)
            Assert.AreEqual(Keys.Control Or Keys.Shift Or Keys.F9, read.Hotkey)
            Assert.AreEqual(ShortcutWindowState.Minimized, read.WindowState)
            TextAssert.PathEquals(workDir, read.WorkingDirectory)
            Assert.IsFalse(read.IsWindowsInstallerShortcut)
            Assert.AreNotEqual(IntPtr.Zero, read.TargetPidl)
            Assert.IsTrue(String.IsNullOrEmpty(read.AppId))

        End Sub

        ''' <summary>
        ''' Verifies that each window state value survives a write/read cycle.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        <DataRow(ShortcutWindowState.Normal)>
        <DataRow(ShortcutWindowState.Maximized)>
        <DataRow(ShortcutWindowState.Minimized)>
        Public Sub Create_WindowState_RoundTrip(windowState As ShortcutWindowState)

            Dim targetPath As String = Me.tempDir.CreateFile("App.txt", "target")
            Dim lnkPath As String = Me.tempDir.Combine("WindowState.lnk")

            Dim written As New ShortcutFileInfo(lnkPath) With {.ViewMode = True}
            written.Target = targetPath
            written.WindowState = windowState
            written.Create()

            Assert.AreEqual(windowState, New ShortcutFileInfo(lnkPath).WindowState)

        End Sub

        ''' <summary>
        ''' Verifies that a shortcut created without assigning a window state is stored as
        ''' <see cref="ShortcutWindowState.Normal"/>, instead of the native SW_HIDE value (0)
        ''' that would make the target launch hidden (regression test).
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub Create_WithoutWindowState_StoresNormal()

            Dim targetPath As String = Me.tempDir.CreateFile("App.txt", "target")
            Dim lnkPath As String = Me.tempDir.Combine("DefaultWindowState.lnk")

            Dim written As New ShortcutFileInfo(lnkPath) With {.ViewMode = True}
            written.Target = targetPath
            written.Create()

            Assert.AreEqual(ShortcutWindowState.Normal, New ShortcutFileInfo(lnkPath).WindowState)

        End Sub

        ''' <summary>
        ''' Verifies that an empty shortcut (as created by the "New" command of the GUI) is stored with
        ''' <see cref="ShortcutWindowState.Normal"/> as window state (regression test).
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub Create_EmptyShortcut_StoresNormalWindowState()

            Dim lnkPath As String = Me.tempDir.Combine("Empty.lnk")

            Dim written As New ShortcutFileInfo(lnkPath) With {.ViewMode = True}
            written.Create()

            Assert.IsTrue(File.Exists(lnkPath))
            Assert.AreEqual(ShortcutWindowState.Normal, New ShortcutFileInfo(lnkPath).WindowState)

        End Sub

        ''' <summary>
        ''' Verifies that a directory can be used as the shortcut target.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub Create_DirectoryTarget_RoundTrip()

            Dim targetDir As String = Me.tempDir.Combine("TargetFolder")
            Directory.CreateDirectory(targetDir)
            Dim lnkPath As String = Me.tempDir.Combine("Folder.lnk")

            Dim written As New ShortcutFileInfo(lnkPath) With {.ViewMode = True}
            written.Target = targetDir
            written.Create()

            TextAssert.PathEquals(targetDir, New ShortcutFileInfo(lnkPath).Target)

        End Sub

        ''' <summary>
        ''' Verifies that the raw IDList stored by the Shell in a real shortcut can be read from its binary content.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub ReadIDListFromLnkBinary_RealShortcut_ReturnsPidl()

            Dim lnkPath As String = Me.CreateShortcut("Binary.lnk")

            Dim pidl As IntPtr = ShortcutFileInfo.ReadIDListFromLnkBinary(lnkPath)
            Try
                Assert.AreNotEqual(IntPtr.Zero, pidl)
            Finally
                Marshal.FreeCoTaskMem(pidl)
            End Try

        End Sub

        ''' <summary>
        ''' Verifies that the PIDL read from a real shortcut can be cloned.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub CloneTargetPidl_RealShortcut_ReturnsCopy()

            Dim shortcut As New ShortcutFileInfo(Me.CreateShortcut("Clone.lnk"))

            Dim clone As IntPtr = shortcut.CloneTargetPidl()
            Try
                Assert.AreNotEqual(IntPtr.Zero, clone)
                Assert.AreNotEqual(shortcut.TargetPidl, clone)
            Finally
                Marshal.FreeCoTaskMem(clone)
            End Try

        End Sub

#End Region

#Region " Immediate Writes Outside ViewMode (Integration) "

        ''' <summary>
        ''' Verifies that, outside view mode, each property assignment is written to disk immediately.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub PropertySetter_NotViewMode_WritesImmediately()

            Dim lnkPath As String = Me.CreateShortcut("Immediate.lnk")
            Dim shortcut As New ShortcutFileInfo(lnkPath)

            shortcut.Description = "Written immediately"
            shortcut.TargetArguments = "--immediate"
            shortcut.Hotkey = Keys.Alt Or Keys.F2

            Dim read As New ShortcutFileInfo(lnkPath)
            Assert.AreEqual("Written immediately", read.Description)
            Assert.AreEqual("--immediate", read.TargetArguments)
            Assert.AreEqual(Keys.Alt Or Keys.F2, read.Hotkey)

        End Sub

        ''' <summary>
        ''' Verifies that <see cref="ShortcutFileInfo.Refresh()"/> reloads properties changed by another instance.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub Refresh_AfterExternalChange_ReloadsProperties()

            Dim lnkPath As String = Me.CreateShortcut("Refresh.lnk")
            Dim observer As New ShortcutFileInfo(lnkPath)
            Dim writer As New ShortcutFileInfo(lnkPath)

            writer.Description = "Changed by another instance"
            Assert.AreEqual("Test shortcut", observer.Description)

            observer.Refresh()

            Assert.AreEqual("Changed by another instance", observer.Description)

        End Sub

#End Region

#Region " Read-Only Attribute (Integration) "

        ''' <summary>
        ''' Verifies that <see cref="ShortcutFileInfo.IsReadOnly"/> can both set and clear the read-only attribute
        ''' (regression test: assigning <see langword="False"/> used to set the attribute too).
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub IsReadOnly_SetTrueThenFalse_TogglesAttribute()

            Dim lnkPath As String = Me.CreateShortcut("ReadOnly.lnk")
            Dim shortcut As New ShortcutFileInfo(lnkPath)
            Assert.IsFalse(shortcut.IsReadOnly)

            shortcut.IsReadOnly = True
            Assert.IsTrue(File.GetAttributes(lnkPath).HasFlag(FileAttributes.ReadOnly))
            Assert.IsTrue(shortcut.IsReadOnly)

            shortcut.IsReadOnly = False
            Assert.IsFalse(File.GetAttributes(lnkPath).HasFlag(FileAttributes.ReadOnly))
            Assert.IsFalse(shortcut.IsReadOnly)

        End Sub

        ''' <summary>
        ''' Verifies that a read-only shortcut can be modified, and that the read-only attribute is restored afterwards.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub PropertySetter_ReadOnlyFile_WritesAndRestoresAttribute()

            Dim lnkPath As String = Me.CreateShortcut("ReadOnlyWrite.lnk")
            File.SetAttributes(lnkPath, File.GetAttributes(lnkPath) Or FileAttributes.ReadOnly)

            Dim shortcut As New ShortcutFileInfo(lnkPath)
            shortcut.Description = "Changed while read-only"

            Assert.IsTrue(File.GetAttributes(lnkPath).HasFlag(FileAttributes.ReadOnly))
            Assert.AreEqual("Changed while read-only", New ShortcutFileInfo(lnkPath).Description)

        End Sub

#End Region

#Region " AppId (Integration) "

        ''' <summary>
        ''' Verifies that the Application User Model ID is written by <see cref="ShortcutFileInfo.Create()"/> and read back.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub AppId_Create_RoundTrip()

            Dim targetPath As String = Me.tempDir.CreateFile("App.txt", "target")
            Dim lnkPath As String = Me.tempDir.Combine("AppId.lnk")
            Const appId As String = "ElektroStudios.EasyLinkFileViewer.Tests.RoundTrip"

            Dim written As New ShortcutFileInfo(lnkPath) With {.ViewMode = True}
            written.Target = targetPath
            written.AppId = appId
            written.Create()

            Assert.AreEqual(appId, New ShortcutFileInfo(lnkPath).AppId)

        End Sub

        ''' <summary>
        ''' Verifies that the Application User Model ID survives subsequent writes of other properties
        ''' (the classic IShellLink save pipeline wipes the property store, so it must be re-injected).
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub AppId_SurvivesSubsequentPropertyWrites()

            Dim targetPath As String = Me.tempDir.CreateFile("App.txt", "target")
            Dim lnkPath As String = Me.tempDir.Combine("AppIdPersist.lnk")
            Const appId As String = "ElektroStudios.EasyLinkFileViewer.Tests.Persist"

            Dim written As New ShortcutFileInfo(lnkPath) With {.ViewMode = True}
            written.Target = targetPath
            written.AppId = appId
            written.Create()

            Dim editor As New ShortcutFileInfo(lnkPath)
            editor.Description = "Another property changed"

            Dim read As New ShortcutFileInfo(lnkPath)
            Assert.AreEqual(appId, read.AppId)
            Assert.AreEqual("Another property changed", read.Description)

        End Sub

        ''' <summary>
        ''' Verifies that assigning an empty Application User Model ID removes it from the shortcut.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub AppId_SetEmpty_RemovesAppId()

            Dim targetPath As String = Me.tempDir.CreateFile("App.txt", "target")
            Dim lnkPath As String = Me.tempDir.Combine("AppIdClear.lnk")

            Dim written As New ShortcutFileInfo(lnkPath) With {.ViewMode = True}
            written.Target = targetPath
            written.AppId = "ElektroStudios.EasyLinkFileViewer.Tests.Clear"
            written.Create()

            Dim editor As New ShortcutFileInfo(lnkPath)
            editor.AppId = String.Empty

            Assert.IsTrue(String.IsNullOrEmpty(ShortcutAppIdHelper.GetAppUserModelId(lnkPath)))

        End Sub

#End Region

#Region " Target Execution Modes (Integration) "

        ''' <summary>
        ''' Verifies that the "relative path via Windows Explorer" mode stores explorer.exe as target
        ''' and the quoted relative path to the real target as arguments.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub Create_RelativePathViaExplorer_StoresExplorerWithRelativeArgument()

            Dim targetPath As String = Me.tempDir.CreateFile("Sub\Portable.txt", "target")
            Dim lnkPath As String = Me.tempDir.Combine("Portable.lnk")
            Dim expectedExplorer As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "explorer.exe")

            Dim shortcut As New ShortcutFileInfo(lnkPath) With {.ViewMode = True}
            shortcut.Target = targetPath
            shortcut.TargetExecutionMode = ShortcutTargetExecutionMode.RelativePathViaExplorer
            shortcut.Create()

            Assert.AreEqual(expectedExplorer, shortcut.Target)
            Assert.AreEqual(""".\Sub\Portable.txt""", shortcut.TargetArguments)

            Dim read As New ShortcutFileInfo(lnkPath)
            TextAssert.PathEquals(expectedExplorer, read.Target)
            Assert.AreEqual(""".\Sub\Portable.txt""", read.TargetArguments)

        End Sub

        ''' <summary>
        ''' Verifies that the "relative path with arguments via Command Prompt" mode stores cmd.exe as target,
        ''' and a "start" command with the relative path followed by the original arguments.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub Create_RelativePathViaCommandPrompt_StoresCmdWithRelativePathAndArguments()

            Dim targetPath As String = Me.tempDir.CreateFile("Sub\Portable.txt", "target")
            Dim lnkPath As String = Me.tempDir.Combine("PortableCmd.lnk")
            Dim expectedCmd As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "cmd.exe")

            Dim shortcut As New ShortcutFileInfo(lnkPath) With {.ViewMode = True}
            shortcut.Target = targetPath
            shortcut.TargetArguments = "--portable"
            shortcut.TargetExecutionMode = ShortcutTargetExecutionMode.RelativePathWithArgumentsViaCommandPrompt
            shortcut.Create()

            Const expectedArguments As String = "/c start """" "".\Sub\Portable.txt"" --portable"
            Assert.AreEqual(expectedCmd, shortcut.Target)
            Assert.AreEqual(expectedArguments, shortcut.TargetArguments)

            Dim read As New ShortcutFileInfo(lnkPath)
            TextAssert.PathEquals(expectedCmd, read.Target)
            Assert.AreEqual(expectedArguments, read.TargetArguments)

        End Sub

        ''' <summary>
        ''' Verifies that the Command Prompt mode without arguments does not append a trailing space.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub Create_RelativePathViaCommandPromptWithoutArguments_StoresOnlyStartCommand()

            Dim targetPath As String = Me.tempDir.CreateFile("Portable.txt", "target")
            Dim lnkPath As String = Me.tempDir.Combine("PortableCmdNoArgs.lnk")

            Dim shortcut As New ShortcutFileInfo(lnkPath) With {.ViewMode = True}
            shortcut.Target = targetPath
            shortcut.TargetExecutionMode = ShortcutTargetExecutionMode.RelativePathWithArgumentsViaCommandPrompt
            shortcut.Create()

            Assert.AreEqual("/c start """" "".\Portable.txt""", shortcut.TargetArguments)

        End Sub

        ''' <summary>
        ''' Verifies that a relative path mode fails with a descriptive exception when the target
        ''' is on a different drive than the shortcut, because no relative path exists between them.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub Create_RelativePathModeAcrossDrives_ThrowsInvalidOperationException()

            Dim tempRoot As String = Path.GetPathRoot(Me.tempDir.FullPath)
            Dim otherDrive As String = If(tempRoot.StartsWith("Q", StringComparison.OrdinalIgnoreCase), "R:\", "Q:\")
            Dim lnkPath As String = Me.tempDir.Combine("CrossDrive.lnk")

            Dim shortcut As New ShortcutFileInfo(lnkPath) With {.ViewMode = True}
            shortcut.Target = $"{otherDrive}Folder\Program.exe"
            shortcut.TargetExecutionMode = ShortcutTargetExecutionMode.RelativePathViaExplorer

            Dim ex As InvalidOperationException = Assert.ThrowsExactly(Of InvalidOperationException)(Sub() shortcut.Create())

            TextAssert.Contains(ex.Message, "relative path")

        End Sub

#End Region

#Region " File Operations (Integration) "

        ''' <summary>
        ''' Verifies that <see cref="ShortcutFileInfo.CopyTo(String, Boolean)"/> copies the shortcut
        ''' and returns an instance that reads the copy.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub CopyTo_CopiesShortcut()

            Dim targetPath As String = Nothing
            Dim lnkPath As String = Me.CreateShortcut("Source.lnk", targetPath)
            Dim copyPath As String = Me.tempDir.Combine("Copy.lnk")

            Dim copy As ShortcutFileInfo = New ShortcutFileInfo(lnkPath).CopyTo(copyPath, overwrite:=False)

            Assert.IsTrue(File.Exists(lnkPath))
            Assert.IsTrue(File.Exists(copyPath))
            Assert.AreEqual(copyPath, copy.FullName)
            TextAssert.PathEquals(targetPath, copy.Target)
            Assert.AreEqual("Test shortcut", copy.Description)

        End Sub

        ''' <summary>
        ''' Verifies that <see cref="ShortcutFileInfo.MoveTo(String)"/> moves the shortcut file.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub MoveTo_MovesShortcutFile()

            Dim lnkPath As String = Me.CreateShortcut("ToMove.lnk")
            Dim movedPath As String = Me.tempDir.Combine("Moved.lnk")

            Dim shortcut As New ShortcutFileInfo(lnkPath)
            shortcut.MoveTo(movedPath)

            Assert.IsFalse(File.Exists(lnkPath))
            Assert.IsTrue(File.Exists(movedPath))
            Assert.AreEqual("Test shortcut", New ShortcutFileInfo(movedPath).Description)

        End Sub

        ''' <summary>
        ''' Verifies that <see cref="ShortcutFileInfo.Delete()"/> deletes the shortcut file.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub Delete_DeletesShortcutFile()

            Dim lnkPath As String = Me.CreateShortcut("ToDelete.lnk")
            Dim shortcut As New ShortcutFileInfo(lnkPath)

            shortcut.Delete()

            Assert.IsFalse(File.Exists(lnkPath))
            Assert.IsFalse(shortcut.Exists)

        End Sub

        ''' <summary>
        ''' Verifies the implicit conversions between <see cref="ShortcutFileInfo"/> and <see cref="FileInfo"/>.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub WideningConversions_FileInfo_RoundTrip()

            Dim lnkPath As String = Me.CreateShortcut("Conversion.lnk")

            Dim convertedFileInfo As FileInfo = New ShortcutFileInfo(lnkPath)
            Dim convertedShortcut As ShortcutFileInfo = New FileInfo(lnkPath)

            Assert.AreEqual(lnkPath, convertedFileInfo.FullName)
            Assert.AreEqual(lnkPath, convertedShortcut.FullName)
            Assert.AreEqual("Test shortcut", convertedShortcut.Description)

        End Sub

        ''' <summary>
        ''' Verifies that <see cref="ShortcutFileInfo.OpenRead()"/> exposes the raw shell link bytes,
        ''' whose header size field is always 0x4C.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub OpenRead_ReturnsShellLinkBinaryContent()

            Dim shortcut As New ShortcutFileInfo(Me.CreateShortcut("Raw.lnk"))
            Dim header As Byte() = New Byte(3) {}

            Using stream As FileStream = shortcut.OpenRead()
                Dim bytesRead As Integer = stream.Read(header, 0, header.Length)
                Assert.AreEqual(header.Length, bytesRead)
            End Using

            Assert.AreEqual(&H4CUI, BitConverter.ToUInt32(header, 0))

        End Sub

#End Region

    End Class

End Namespace

#End Region
