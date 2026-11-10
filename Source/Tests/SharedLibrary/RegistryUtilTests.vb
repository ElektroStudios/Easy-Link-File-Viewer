#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports Microsoft.VisualStudio.TestTools.UnitTesting

#End Region

#Region " RegistryUtilTests "

Namespace EasyLink.Tests.SharedLibrary

    ''' <summary>
    ''' Unit tests for the argument validation of the <see cref="RegistryUtil"/> module.
    ''' <para></para>
    ''' The registry itself is never touched: writing to HKEY_CLASSES_ROOT requires administrator rights
    ''' and would alter the machine configuration, so only the fail-fast preconditions are verified.
    ''' </summary>
    <TestClass>
    Public NotInheritable Class RegistryUtilTests

#Region " CreateFileTypeRegistryMenuEntry "

        ''' <summary>
        ''' Verifies that a missing file type is rejected.
        ''' </summary>
        <TestMethod>
        <DataRow("")>
        <DataRow("   ")>
        Public Sub CreateEntry_EmptyFileType_ThrowsArgumentNullException(fileType As String)

            Dim ex As ArgumentNullException = Assert.ThrowsExactly(Of ArgumentNullException)(
                Sub() RegistryUtil.CreateFileTypeRegistryMenuEntry(fileType, "key", "text", "", "icon", "command"))

            Assert.AreEqual("fileType", ex.ParamName)

        End Sub

        ''' <summary>
        ''' Verifies that a missing key name is rejected.
        ''' </summary>
        <TestMethod>
        <DataRow("")>
        <DataRow("   ")>
        Public Sub CreateEntry_EmptyKeyName_ThrowsArgumentNullException(keyName As String)

            Dim ex As ArgumentNullException = Assert.ThrowsExactly(Of ArgumentNullException)(
                Sub() RegistryUtil.CreateFileTypeRegistryMenuEntry("lnkfile", keyName, "text", "", "icon", "command"))

            Assert.AreEqual("keyName", ex.ParamName)

        End Sub

        ''' <summary>
        ''' Verifies that a missing command is rejected.
        ''' </summary>
        <TestMethod>
        <DataRow("")>
        <DataRow("   ")>
        Public Sub CreateEntry_EmptyCommand_ThrowsArgumentNullException(command As String)

            Dim ex As ArgumentNullException = Assert.ThrowsExactly(Of ArgumentNullException)(
                Sub() RegistryUtil.CreateFileTypeRegistryMenuEntry("lnkfile", "key", "text", "", "icon", command))

            Assert.AreEqual("command", ex.ParamName)

        End Sub

        ''' <summary>
        ''' Verifies that <see langword="Nothing"/> required arguments are rejected.
        ''' </summary>
        <TestMethod>
        Public Sub CreateEntry_NothingArguments_ThrowsArgumentNullException()

            Assert.ThrowsExactly(Of ArgumentNullException)(
                Sub() RegistryUtil.CreateFileTypeRegistryMenuEntry(Nothing, "key", "text", "", "icon", "command"))
            Assert.ThrowsExactly(Of ArgumentNullException)(
                Sub() RegistryUtil.CreateFileTypeRegistryMenuEntry("lnkfile", Nothing, "text", "", "icon", "command"))
            Assert.ThrowsExactly(Of ArgumentNullException)(
                Sub() RegistryUtil.CreateFileTypeRegistryMenuEntry("lnkfile", "key", "text", "", "icon", Nothing))

        End Sub

#End Region

#Region " DeleteFileTypeRegistryMenuEntry "

        ''' <summary>
        ''' Verifies that a missing file type is rejected.
        ''' </summary>
        <TestMethod>
        <DataRow("")>
        <DataRow("   ")>
        Public Sub DeleteEntry_EmptyFileType_ThrowsArgumentNullException(fileType As String)

            Dim ex As ArgumentNullException = Assert.ThrowsExactly(Of ArgumentNullException)(
                Sub() RegistryUtil.DeleteFileTypeRegistryMenuEntry(fileType, "key"))

            Assert.AreEqual("fileType", ex.ParamName)

        End Sub

        ''' <summary>
        ''' Verifies that a missing key name is rejected.
        ''' </summary>
        <TestMethod>
        <DataRow("")>
        <DataRow("   ")>
        Public Sub DeleteEntry_EmptyKeyName_ThrowsArgumentNullException(keyName As String)

            Dim ex As ArgumentNullException = Assert.ThrowsExactly(Of ArgumentNullException)(
                Sub() RegistryUtil.DeleteFileTypeRegistryMenuEntry("lnkfile", keyName))

            Assert.AreEqual("keyName", ex.ParamName)

        End Sub

#End Region

    End Class

End Namespace

#End Region
