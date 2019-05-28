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
        Public Shared Function ExtractIconFromFile(ByVal filepath As String) As Icon

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

            Dim ico As Icon = TryCast(Drawing.Icon.FromHandle(shInfo.IconHandle).Clone(), Drawing.Icon)
            NativeMethods.DestroyIcon(shInfo.IconHandle)
            Return ico

        End Function

#End Region

    End Class

End Namespace

#End Region
