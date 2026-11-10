#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.Collections.Generic
Imports System.IO

Imports DevCase.Core.IO

Imports EasyLink.Tests.Helpers

Imports Microsoft.VisualStudio.TestTools.UnitTesting

#End Region

#Region " ShortcutAppIdHelperTests "

Namespace EasyLink.Tests.SharedLibrary

    ''' <summary>
    ''' Unit and integration tests for the <see cref="ShortcutAppIdHelper"/> and <see cref="InstalledAppInfo"/> classes.
    ''' </summary>
    <STATestClass>
    Public NotInheritable Class ShortcutAppIdHelperTests

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
        ''' Creates a real shortcut file in the temporary directory, without an Application User Model ID.
        ''' </summary>
        '''
        ''' <returns>
        ''' The full path of the created shortcut file.
        ''' </returns>
        Private Function CreateShortcut() As String

            Dim targetPath As String = Me.tempDir.CreateFile("App.txt", "target")
            Dim lnkPath As String = Me.tempDir.Combine("AppIdHelper.lnk")

            Dim shortcut As New ShortcutFileInfo(lnkPath) With {.ViewMode = True}
            shortcut.Target = targetPath
            shortcut.Create()

            Return lnkPath

        End Function

#End Region

#Region " Argument Validation "

        ''' <summary>
        ''' Verifies that <see cref="ShortcutAppIdHelper.GetAppUserModelId(String)"/> rejects null or empty paths.
        ''' </summary>
        <TestMethod>
        Public Sub GetAppUserModelId_NullOrEmptyPath_ThrowsArgumentNullException()

            Assert.ThrowsExactly(Of ArgumentNullException)(Sub() ShortcutAppIdHelper.GetAppUserModelId(Nothing))
            Assert.ThrowsExactly(Of ArgumentNullException)(Sub() ShortcutAppIdHelper.GetAppUserModelId(String.Empty))

        End Sub

        ''' <summary>
        ''' Verifies that <see cref="ShortcutAppIdHelper.GetAppUserModelId(String)"/> rejects missing files.
        ''' </summary>
        <TestMethod>
        Public Sub GetAppUserModelId_MissingFile_ThrowsFileNotFoundException()

            Dim lnkPath As String = Me.tempDir.Combine("Missing.lnk")

            Assert.ThrowsExactly(Of FileNotFoundException)(Sub() ShortcutAppIdHelper.GetAppUserModelId(lnkPath))

        End Sub

        ''' <summary>
        ''' Verifies that <see cref="ShortcutAppIdHelper.SetAppUserModelId(String, String)"/> rejects null or empty paths.
        ''' </summary>
        <TestMethod>
        Public Sub SetAppUserModelId_NullOrEmptyPath_ThrowsArgumentNullException()

            Assert.ThrowsExactly(Of ArgumentNullException)(Sub() ShortcutAppIdHelper.SetAppUserModelId(Nothing, "Company.Product"))
            Assert.ThrowsExactly(Of ArgumentNullException)(Sub() ShortcutAppIdHelper.SetAppUserModelId(String.Empty, "Company.Product"))

        End Sub

        ''' <summary>
        ''' Verifies that <see cref="ShortcutAppIdHelper.SetAppUserModelId(String, String)"/> rejects a null AppId.
        ''' </summary>
        <TestMethod>
        Public Sub SetAppUserModelId_NullAppId_ThrowsArgumentNullException()

            Dim lnkPath As String = Me.tempDir.Combine("Missing.lnk")

            Dim ex As ArgumentNullException = Assert.ThrowsExactly(Of ArgumentNullException)(
                Sub() ShortcutAppIdHelper.SetAppUserModelId(lnkPath, Nothing))

            Assert.AreEqual("appId", ex.ParamName)

        End Sub

        ''' <summary>
        ''' Verifies that <see cref="ShortcutAppIdHelper.SetAppUserModelId(String, String)"/> rejects missing files.
        ''' </summary>
        <TestMethod>
        Public Sub SetAppUserModelId_MissingFile_ThrowsFileNotFoundException()

            Dim lnkPath As String = Me.tempDir.Combine("Missing.lnk")

            Assert.ThrowsExactly(Of FileNotFoundException)(Sub() ShortcutAppIdHelper.SetAppUserModelId(lnkPath, "Company.Product"))

        End Sub

        ''' <summary>
        ''' Verifies that <see cref="ShortcutAppIdHelper.ClearAppUserModelId(String)"/> rejects null or empty paths.
        ''' </summary>
        <TestMethod>
        Public Sub ClearAppUserModelId_NullOrEmptyPath_ThrowsArgumentNullException()

            Assert.ThrowsExactly(Of ArgumentNullException)(Sub() ShortcutAppIdHelper.ClearAppUserModelId(Nothing))
            Assert.ThrowsExactly(Of ArgumentNullException)(Sub() ShortcutAppIdHelper.ClearAppUserModelId(String.Empty))

        End Sub

        ''' <summary>
        ''' Verifies that <see cref="ShortcutAppIdHelper.ClearAppUserModelId(String)"/> rejects missing files.
        ''' </summary>
        <TestMethod>
        Public Sub ClearAppUserModelId_MissingFile_ThrowsFileNotFoundException()

            Dim lnkPath As String = Me.tempDir.Combine("Missing.lnk")

            Assert.ThrowsExactly(Of FileNotFoundException)(Sub() ShortcutAppIdHelper.ClearAppUserModelId(lnkPath))

        End Sub

#End Region

#Region " Set / Get / Clear (Integration) "

        ''' <summary>
        ''' Verifies that a shortcut without an Application User Model ID returns <see langword="Nothing"/>.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub GetAppUserModelId_ShortcutWithoutAppId_ReturnsNullOrEmpty()

            Dim lnkPath As String = Me.CreateShortcut()

            Assert.IsTrue(String.IsNullOrEmpty(ShortcutAppIdHelper.GetAppUserModelId(lnkPath)))

        End Sub

        ''' <summary>
        ''' Verifies the set, get, overwrite and clear cycle of the Application User Model ID.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub SetGetClear_FullCycle()

            Dim lnkPath As String = Me.CreateShortcut()

            ShortcutAppIdHelper.SetAppUserModelId(lnkPath, "ElektroStudios.Tests.First")
            Assert.AreEqual("ElektroStudios.Tests.First", ShortcutAppIdHelper.GetAppUserModelId(lnkPath))

            ShortcutAppIdHelper.SetAppUserModelId(lnkPath, "ElektroStudios.Tests.Second")
            Assert.AreEqual("ElektroStudios.Tests.Second", ShortcutAppIdHelper.GetAppUserModelId(lnkPath))

            ShortcutAppIdHelper.ClearAppUserModelId(lnkPath)
            Assert.IsTrue(String.IsNullOrEmpty(ShortcutAppIdHelper.GetAppUserModelId(lnkPath)))

        End Sub

        ''' <summary>
        ''' Verifies that setting the Application User Model ID does not alter the rest of the shortcut properties.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub SetAppUserModelId_KeepsOtherShortcutProperties()

            Dim lnkPath As String = Me.CreateShortcut()
            Dim targetBefore As String = New ShortcutFileInfo(lnkPath).Target

            ShortcutAppIdHelper.SetAppUserModelId(lnkPath, "ElektroStudios.Tests.KeepProperties")

            Dim shortcut As New ShortcutFileInfo(lnkPath)
            Assert.AreEqual(targetBefore, shortcut.Target)
            Assert.AreEqual("ElektroStudios.Tests.KeepProperties", shortcut.AppId)

        End Sub

        ''' <summary>
        ''' Verifies that enumerating the installed applications returns a list
        ''' in which every entry has an Application User Model ID.
        ''' </summary>
        <TestMethod>
        <TestCategory("Integration")>
        Public Sub GetInstalledAppUserModelIds_ReturnsEntriesWithAppIds()

            Dim apps As List(Of InstalledAppInfo) = ShortcutAppIdHelper.GetInstalledAppUserModelIds()

            Assert.IsNotNull(apps)
            For Each app As InstalledAppInfo In apps
                Assert.IsFalse(String.IsNullOrEmpty(app.AppUserModelId), $"Entry without AppUserModelId: {app}")
                Assert.IsNotNull(app.DisplayName)
            Next app

        End Sub

#End Region

#Region " InstalledAppInfo "

        ''' <summary>
        ''' Verifies the properties and string representation of <see cref="InstalledAppInfo"/>.
        ''' </summary>
        <TestMethod>
        Public Sub InstalledAppInfo_PropertiesAndToString()

            Dim app As New InstalledAppInfo("Calculator", "Microsoft.WindowsCalculator_8wekyb3d8bbwe!App")

            Assert.AreEqual("Calculator", app.DisplayName)
            Assert.AreEqual("Microsoft.WindowsCalculator_8wekyb3d8bbwe!App", app.AppUserModelId)
            Assert.AreEqual("Calculator (Microsoft.WindowsCalculator_8wekyb3d8bbwe!App)", app.ToString())

        End Sub

#End Region

    End Class

End Namespace

#End Region
