#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.Runtime.InteropServices
Imports System.Security
Imports System.Text

#End Region

#Region " NativeMethods "

Namespace EasyLink.Tests.Helpers

    ''' <summary>
    ''' Platform Invocation (P/Invoke) declarations used by the test helpers.
    ''' </summary>
    <SuppressUnmanagedCodeSecurity>
    Friend NotInheritable Class NativeMethods

#Region " Constructors "

        ''' <summary>
        ''' Prevents a default instance of the <see cref="NativeMethods"/> class from being created.
        ''' </summary>
        Private Sub New()
        End Sub

#End Region

#Region " Kernel32.dll "

        ''' <summary>
        ''' Converts the specified path to its long form.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-getlongpathnamew"/>
        ''' </remarks>
        '''
        ''' <param name="shortPath">
        ''' The path to be converted.
        ''' </param>
        '''
        ''' <param name="longPath">
        ''' A buffer that receives the long path.
        ''' </param>
        '''
        ''' <param name="bufferLength">
        ''' The size of the buffer, in characters.
        ''' </param>
        '''
        ''' <returns>
        ''' If the function succeeds, the length of the string copied to <paramref name="longPath"/>;
        ''' if the buffer is too small, the required buffer size; if the function fails, zero.
        ''' </returns>
        <DllImport("kernel32.dll", EntryPoint:="GetLongPathNameW", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Friend Shared Function GetLongPathName(shortPath As String,
                                               longPath As StringBuilder,
                                               bufferLength As Integer) As Integer
        End Function

#End Region

    End Class

End Namespace

#End Region
