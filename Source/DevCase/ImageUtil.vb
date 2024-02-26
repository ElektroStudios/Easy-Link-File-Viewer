' This source-code is freely distributed as part of "DevCase for .NET Framework".
' 
' Maybe you would like to consider to buy this powerful set of libraries to support me. 
' You can do loads of things with my apis for a big amount of diverse thematics.
' 
' Here is a link to the purchase page:
' https://codecanyon.net/item/elektrokit-class-library-for-net/19260282
' 
' Thank you.

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.IO
Imports System.Runtime.InteropServices

Imports DevCase.Interop.Unmanaged.Win32
Imports DevCase.Interop.Unmanaged.Win32.Enums
Imports DevCase.Interop.Unmanaged.Win32.Structures

#End Region

#Region " Image Util "

Namespace DevCase.Core.Imaging.Tools

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains image related utilities.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------   
    Friend NotInheritable Class ImageUtil

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Prevents a default instance of the <see cref="ImageUtil"/> class from being created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerNonUserCode>
        Private Sub New()
        End Sub

#End Region

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Extracts the icon associated for the specified directory.
        ''' <para></para>
        ''' Note: the maximum size of the returned icon only can be 32x32.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Dim path As String = "C:\Windows"
        ''' Dim ico As Icon = ExtractIconFromDirectory(path)
        ''' Dim bmp As Bitmap = ico.ToBitmap()
        ''' PictureBox1.BackgroundImage = bmp
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="directoryPath">
        ''' The full path to a directory.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The resulting icon.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function ExtractIconFromDirectory(directoryPath As String) As Icon

            ' Note that the icon returned by "SHGetFileInfo" function 
            ' is limited to "SHGetFileInfoFlags.IconSizeSmall" (16x16) 
            ' and "SHGetFileInfoFlags.IconSizeLarge" (32x32) icon size.

            Dim shInfo As New ShellFileInfo()

            Dim result As IntPtr = NativeMethods.SHGetFileInfo(directoryPath, FileAttributes.Directory, shInfo, CUInt(Marshal.SizeOf(shInfo)), SHGetFileInfoFlags.Icon Or SHGetFileInfoFlags.IconSizeLarge)

            If (result = IntPtr.Zero) Then
                Return Nothing
            End If

            Dim ico As Icon = TryCast(Icon.FromHandle(shInfo.IconHandle).Clone(), Icon)
            NativeMethods.DestroyIcon(shInfo.IconHandle)
            Return ico

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Extracts the icon associated for the specified file.
        ''' <para></para>
        ''' Note: the maximum size of the returned icon only can be 32x32.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Dim path As String = "C:\File.ext"
        ''' Dim ico As Icon = ExtractIconFromFile(path)
        ''' Dim bmp As Bitmap = ico.ToBitmap()
        ''' PictureBox1.BackgroundImage = bmp
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="filepath">
        ''' The full path to a file.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The resulting icon.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function ExtractIconFromFile(filepath As String) As Icon

            ' Note that the icon returned by "SHGetFileInfo" function 
            ' is limited to "SHGetFileInfoFlags.IconSizeSmall" (16x16) 
            ' and "SHGetFileInfoFlags.IconSizeLarge" (32x32) icon size.

            Dim shInfo As New ShellFileInfo()

            Dim result As IntPtr = NativeMethods.SHGetFileInfo(filepath, FileAttributes.Normal, shInfo, CUInt(Marshal.SizeOf(shInfo)), SHGetFileInfoFlags.UseFileAttributes Or SHGetFileInfoFlags.Icon Or SHGetFileInfoFlags.IconSizeLarge)
            If (result = IntPtr.Zero) Then
                result = NativeMethods.SHGetFileInfo(filepath, FileAttributes.Normal, shInfo, CUInt(Marshal.SizeOf(shInfo)), SHGetFileInfoFlags.Icon Or SHGetFileInfoFlags.IconSizeLarge)
            End If

            If (result = IntPtr.Zero) Then
                Return Nothing
            End If

            Dim ico As Icon = TryCast(Icon.FromHandle(shInfo.IconHandle).Clone(), Icon)
            NativeMethods.DestroyIcon(shInfo.IconHandle)
            Return ico

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Extracts the icon associated for the specified file extension.
        ''' <para></para>
        ''' Note: the maximum size of the returned icon only can be 32x32.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Dim ext As String = ".txt"
        ''' Dim ico As Icon = ExtractIconFromFileExtension(ext)
        ''' Dim bmp As Bitmap = ico.ToBitmap()
        ''' PictureBox1.BackgroundImage = bmp
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="ext">
        ''' The file extension.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The resulting icon.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function ExtractIconFromFileExtension(ext As String) As Icon

            ' Note that the icon returned by "SHGetFileInfo" function 
            ' is limited to "SHGetFileInfoFlags.IconSizeSmall" (16x16) 
            ' and "SHGetFileInfoFlags.IconSizeLarge" (32x32) icon size.

            Dim shInfo As New ShellFileInfo()

            Dim result As IntPtr = NativeMethods.SHGetFileInfo(ext, FileAttributes.Normal, shInfo, CUInt(Marshal.SizeOf(shInfo)), SHGetFileInfoFlags.UseFileAttributes Or SHGetFileInfoFlags.Icon Or SHGetFileInfoFlags.IconSizeLarge)
            If (result = IntPtr.Zero) Then
                result = NativeMethods.SHGetFileInfo(ext, FileAttributes.Normal, shInfo, CUInt(Marshal.SizeOf(shInfo)), SHGetFileInfoFlags.Icon Or SHGetFileInfoFlags.IconSizeLarge)
            End If

            If (result = IntPtr.Zero) Then
                Return Nothing
            End If

            Dim ico As Icon = TryCast(Icon.FromHandle(shInfo.IconHandle).Clone(), Icon)
            NativeMethods.DestroyIcon(shInfo.IconHandle)
            Return ico

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Extracts a icon stored in the specified executable, dll or icon file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Dim ico As Icon = ExtractIconFromExecutableFile("C:\Windows\Explorer.exe", 0)
        ''' Dim bmp As Bitmap = ico.ToBitmap()
        ''' PictureBox1.BackgroundImage = bmp
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="filepath">
        ''' The source filepath.
        ''' </param>
        ''' 
        ''' <param name="iconIndex">
        ''' The index of the icon to be extracted.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function ExtractIconFromExecutableFile(filepath As String,
                                                             iconIndex As Integer) As Icon

            Return ImageUtil.ExtractIconFromExecutableFile(filepath, iconIndex, 0)

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Extracts a icon stored in the specified executable, dll or icon file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Dim ico As Icon = ExtractIconFromExecutableFile("C:\Windows\Explorer.exe", 0, 256)
        ''' Dim bmp As Bitmap = ico.ToBitmap()
        ''' PictureBox1.BackgroundImage = bmp
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="filepath">
        ''' The source filepath.
        ''' </param>
        ''' 
        ''' <param name="iconIndex">
        ''' The index of the icon to be extracted.
        ''' </param>
        ''' 
        ''' <param name="iconSize">
        ''' The icon size, in pixels.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function ExtractIconFromExecutableFile(filepath As String,
                                                             iconIndex As Integer,
                                                             iconSize As Integer) As Icon

            Dim hiconLarge As IntPtr
            Dim result As Integer = NativeMethods.SHDefExtractIcon(filepath, iconIndex, 0, hiconLarge, Nothing, CUInt(iconSize))

            If (CType(result, HResult) <> HResult.S_OK) Then
                Marshal.ThrowExceptionForHR(result)
                Return Nothing

            Else
                Dim ico As Icon = TryCast(Icon.FromHandle(hiconLarge).Clone(), Icon)
                NativeMethods.DestroyIcon(hiconLarge)
                Return ico

            End If

        End Function

#End Region

    End Class

End Namespace

#End Region
