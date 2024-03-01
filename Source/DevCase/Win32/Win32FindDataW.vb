' This source-code is freely distributed as part of "DevCase Class Library .NET Developers".
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

Imports System.Runtime.InteropServices

#End Region

#Region " Win32FindData (Unicode) "

Namespace DevCase.Interop.Unmanaged.Win32.Structures

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains information about the file that is found by the FindFirstFile, FindFirstFileEx, or FindNextFile function.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/aa365740%28v=vs.85%29.aspx"/>
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Friend Structure Win32FindDataW

        ''' <summary>
        ''' The FILE_ATTRIBUTE_SPARSE_FILE attribute on the file is set 
        ''' if any of the streams of the file have ever been sparse.
        ''' </summary>
        Public FileAttributes As UInteger

        ''' <summary>
        ''' A FILETIME structure that specifies when a file or directory was created.
        ''' If the underlying file system does not support creation time, this member is zero.
        ''' </summary>
        Public CreationTime As Long

        ''' <summary>
        ''' A FILETIME structure.
        ''' <para></para>
        ''' For a file, the structure specifies when the file was last read from, written to, or for executable files, run.
        ''' <para></para>
        ''' For a directory, the structure specifies when the directory is created. 
        ''' <para></para>
        ''' If the underlying file system does not support last access time, this member is zero.
        ''' </summary>
        Public LastAccessTime As Long

        ''' <summary>
        ''' A FILETIME structure.
        ''' <para></para>
        ''' For a file, the structure specifies when the file was last written to, truncated, or overwritten, 
        ''' for example, when WriteFile or SetEndOfFile are used.
        ''' <para></para>
        ''' The date and time are not updated when file attributes or security descriptors are changed.
        ''' </summary>
        Public LastWriteTime As Long

        ''' <summary>
        ''' The high-order DWORD value of the file size, in bytes.
        ''' <para></para>
        ''' This value is zero unless the file size is greater than MAXDWORD.
        ''' <para></para>
        ''' The size of the file is equal to (FileSizeHigh * (MAXDWORD+1)) + FileSizeLow.
        ''' </summary>
        Public FileSizeHigh As UInteger

        ''' <summary>
        ''' The low-order DWORD value of the file size, in bytes.
        ''' </summary>
        Public FileSizeLow As UInteger

        ''' <summary>
        ''' If the <see cref="FileAttributes"/> member includes the FILE_ATTRIBUTE_REPARSE_POINT attribute, 
        ''' this member specifies the reparse point tag.
        ''' Otherwise, this value is undefined and should not be used.
        ''' </summary>
        Public Reserved0 As UInteger

        ''' <summary>
        ''' Reserved for future use.
        ''' </summary>
        Public Reserved1 As UInteger

        ''' <summary>
        ''' The name of the file.
        ''' </summary>
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public FileName As String

        ''' <summary>
        ''' An alternative name for the file.
        ''' This name is in the classic 8.3 file name format.
        ''' </summary>
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=14)>
        Public AlternateFileName As String

    End Structure

End Namespace

#End Region
