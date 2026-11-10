#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading

#End Region

#Region " TempDirectory "

Namespace EasyLink.Tests.Helpers

    ''' <summary>
    ''' Creates a unique, empty temporary directory for a test and deletes it (with all its contents) when disposed.
    ''' <para></para>
    ''' The directory path is always returned in its long (non 8.3) form, so it can be compared
    ''' against paths resolved by the Windows Shell when reading back shortcut targets.
    ''' </summary>
    Friend NotInheritable Class TempDirectory : Implements IDisposable

#Region " Private Fields "

        ''' <summary>
        ''' The maximum number of attempts to delete the directory on dispose.
        ''' </summary>
        Private Const MaxDeleteAttempts As Integer = 10

        ''' <summary>
        ''' The delay, in milliseconds, between delete attempts.
        ''' </summary>
        Private Const DeleteRetryDelayMs As Integer = 100

        ''' <summary>
        ''' Flag to detect redundant calls to <see cref="TempDirectory.Dispose()"/>.
        ''' </summary>
        Private isDisposed As Boolean

#End Region

#Region " Properties "

        ''' <summary>
        ''' Gets the full path of the temporary directory.
        ''' </summary>
        Friend ReadOnly Property FullPath As String

#End Region

#Region " Constructors "

        ''' <summary>
        ''' Initializes a new instance of the <see cref="TempDirectory"/> class,
        ''' creating a new unique directory inside the user's temporary folder.
        ''' </summary>
        Friend Sub New()

            Dim baseTempPath As String = TempDirectory.GetLongPath(Path.GetTempPath())
            Dim directoryName As String = $"EasyLink.Tests_{Guid.NewGuid():N}"
            Me.FullPath = Path.Combine(baseTempPath, directoryName)

            Directory.CreateDirectory(Me.FullPath)

        End Sub

#End Region

#Region " Public Methods "

        ''' <summary>
        ''' Combines the temporary directory path with the specified relative path.
        ''' </summary>
        '''
        ''' <param name="relativePath">
        ''' The relative path to combine.
        ''' </param>
        '''
        ''' <returns>
        ''' The combined full path.
        ''' </returns>
        Friend Function Combine(relativePath As String) As String

            If String.IsNullOrWhiteSpace(relativePath) Then
                Throw New ArgumentNullException(NameOf(relativePath))
            End If

            Return Path.Combine(Me.FullPath, relativePath)

        End Function

        ''' <summary>
        ''' Creates a file with the specified relative path and text content inside the temporary directory,
        ''' creating any missing parent directories.
        ''' </summary>
        '''
        ''' <param name="relativePath">
        ''' The relative path of the file to create.
        ''' </param>
        '''
        ''' <param name="content">
        ''' The text content to write.
        ''' </param>
        '''
        ''' <returns>
        ''' The full path of the created file.
        ''' </returns>
        Friend Function CreateFile(relativePath As String, content As String) As String

            Dim filePath As String = Me.Combine(relativePath)
            Directory.CreateDirectory(Path.GetDirectoryName(filePath))
            File.WriteAllText(filePath, content, Encoding.UTF8)
            Return filePath

        End Function

        ''' <summary>
        ''' Creates a file with the specified relative path and binary content inside the temporary directory.
        ''' </summary>
        '''
        ''' <param name="relativePath">
        ''' The relative path of the file to create.
        ''' </param>
        '''
        ''' <param name="content">
        ''' The binary content to write.
        ''' </param>
        '''
        ''' <returns>
        ''' The full path of the created file.
        ''' </returns>
        Friend Function CreateBinaryFile(relativePath As String, content As Byte()) As String

            If content Is Nothing Then
                Throw New ArgumentNullException(NameOf(content))
            End If

            Dim filePath As String = Me.Combine(relativePath)
            Directory.CreateDirectory(Path.GetDirectoryName(filePath))
            File.WriteAllBytes(filePath, content)
            Return filePath

        End Function

#End Region

#Region " Private Methods "

        ''' <summary>
        ''' Converts the specified path to its long (non 8.3) form.
        ''' </summary>
        '''
        ''' <param name="pathToConvert">
        ''' The path to convert. It must exist.
        ''' </param>
        '''
        ''' <returns>
        ''' The long form of the path.
        ''' </returns>
        '''
        ''' <exception cref="Win32Exception">
        ''' Thrown when the native function fails.
        ''' </exception>
        Private Shared Function GetLongPath(pathToConvert As String) As String

            Dim buffer As New StringBuilder(capacity:=260)
            Dim length As Integer = NativeMethods.GetLongPathName(pathToConvert, buffer, buffer.Capacity)

            If length > buffer.Capacity Then
                buffer = New StringBuilder(capacity:=length)
                length = NativeMethods.GetLongPathName(pathToConvert, buffer, buffer.Capacity)
            End If

            If length = 0 Then
                Throw New Win32Exception(Marshal.GetLastWin32Error(), $"GetLongPathName failed for path: {pathToConvert}")
            End If

            Return buffer.ToString()

        End Function

        ''' <summary>
        ''' Removes the read-only attribute from every file inside the temporary directory,
        ''' so the directory tree can be deleted.
        ''' </summary>
        Private Sub ClearReadOnlyAttributes()

            For Each filePath As String In Directory.EnumerateFiles(Me.FullPath, "*", SearchOption.AllDirectories)
                Thread.CurrentThread.Join(0)

                Dim attributes As FileAttributes = File.GetAttributes(filePath)
                If attributes.HasFlag(FileAttributes.ReadOnly) Then
                    File.SetAttributes(filePath, attributes And Not FileAttributes.ReadOnly)
                End If
            Next filePath

        End Sub

#End Region

#Region " IDisposable Implementation "

        ''' <summary>
        ''' Deletes the temporary directory and all its contents.
        ''' <para></para>
        ''' The Windows Shell may briefly keep a handle open to recently written shortcut files,
        ''' so the deletion is retried a few times before giving up with the last exception.
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose

            If Me.isDisposed Then
                Exit Sub
            End If

            If Directory.Exists(Me.FullPath) Then
                Me.ClearReadOnlyAttributes()

                Dim attempt As Integer = 1
                Do
                    Try
                        Directory.Delete(Me.FullPath, recursive:=True)
                        Exit Do

                    Catch ex As IOException When attempt < MaxDeleteAttempts
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                        Thread.Sleep(DeleteRetryDelayMs)
                        attempt += 1

                    Catch ex As UnauthorizedAccessException When attempt < MaxDeleteAttempts
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                        Thread.Sleep(DeleteRetryDelayMs)
                        attempt += 1

                    End Try
                Loop
            End If

            Me.isDisposed = True

        End Sub

#End Region

    End Class

End Namespace

#End Region
