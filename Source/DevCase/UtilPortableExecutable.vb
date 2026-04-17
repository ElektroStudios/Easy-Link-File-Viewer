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

Imports System.ComponentModel
Imports System.IO
Imports System.Reflection.Emit
Imports System.Reflection.PortableExecutable

#End Region

#Region " UtilPortableExecutable "

' ReSharper disable once CheckNamespace

Namespace DevCase.Core.Diagnostics.PortableExecutable

    ''' <summary>
    ''' Contains PE header related utilities.
    ''' </summary>
    <ImmutableObject(True)>
    Public NotInheritable Class UtilPortableExecutable

#Region " Constructors "

        ''' <summary>
        ''' Prevents a default instance of the <see cref="UtilPortableExecutable"/> class from being created.
        ''' </summary>
        <DebuggerNonUserCode>
        Private Sub New()
        End Sub

#End Region

#Region " Public Methods "

        ''' <summary>
        ''' Gets the <see cref="PEFileKinds"/> of the specified Portable Executable (PE) file.
        ''' </summary>
        '''
        ''' <remarks>
        ''' Note: To use this functionality you need to install this nuget package:
        ''' <para></para>
        ''' <see href="https://www.nuget.org/packages/System.Reflection.Metadata">System.Reflection.Metadata</see>
        ''' </remarks>
        '''
        ''' <example> This is a code example.
        ''' <code language="VB.NET">
        ''' Dim filePath As String = "C:\MyExecutable.exe"
        ''' Dim peFileKind As PEFileKinds = GetPEFileKind(filePath)
        ''' 
        ''' Console.WriteLine(peFileKind.ToString())
        ''' </code>
        ''' </example>
        '''
        ''' <param name="filepath">
        ''' Path to the Portable Executable (PE) file.
        ''' </param>
        '''
        ''' <returns>
        ''' The resulting <see cref="PEFileKinds"/>.
        ''' </returns>
        <DebuggerStepThrough>
        Public Shared Function GetPEFileKind(filepath As String) As PEFileKinds

            Const BufferSizeFileStreamDefault As Integer = 4096
            Using fs As New FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.Read,
                                       BufferSizeFileStreamDefault, FileOptions.None),
                  peReader As New PEReader(fs, PEStreamOptions.Default)

                Dim headers As PEHeaders = peReader.PEHeaders

                ' NOTE: the header values cause 'PEHeaders.IsConsoleApplication' to return true for DLLs, 
                ' so we need to check 'PEHeaders.IsDll' first.
                If headers.IsDll Then
                    Return PEFileKinds.Dll

                ElseIf headers.IsConsoleApplication Then
                    Return PEFileKinds.ConsoleApplication

                Else
                    Return PEFileKinds.WindowApplication

                End If

            End Using

        End Function

        ''' <summary>
        ''' Gets the <see cref="PEFileKinds"/> of the specified .NET assembly.
        ''' </summary>
        '''
        ''' <remarks>
        ''' Note: To use this functionality you need to install this nuget package:
        ''' <para></para>
        ''' <see href="https://www.nuget.org/packages/System.Reflection.Metadata">System.Reflection.Metadata</see>
        ''' </remarks>
        '''
        ''' <param name="assembly">
        ''' The .NET assembly.
        ''' </param>
        '''
        ''' <returns>
        ''' The resulting <see cref="PEFileKinds"/>.
        ''' </returns>
        <DebuggerStepThrough>
        Public Shared Function GetPEFileKind(assembly As System.Reflection.Assembly) As PEFileKinds

            Dim assemblyPath As String = assembly.Location
            If String.IsNullOrEmpty(assemblyPath) Then
                Throw New InvalidOperationException("The assembly does not have a physical location.")
            End If

            Return UtilPortableExecutable.GetPEFileKind(assemblyPath)
        End Function

#End Region

    End Class

End Namespace

#If NETCOREAPP Then

Namespace DevCase.Core.Diagnostics.PortableExecutable

    ' The PEFileKinds enumeration is not present in the NuGet package of 'System.Reflection.Emit' for .NET Core,
    ' so we recreate it for this purpose.

    ''' <summary>
    ''' Specifies the type of the portable executable (PE) file.
    ''' </summary>
    <Serializable>
    <ComVisible(True)>
    Public Enum PEFileKinds

        ''' <summary>
        ''' The portable executable (PE) file is a DLL.
        ''' </summary>
        Dll = 1

        ''' <summary>
        ''' The application is a console (not a Windows-based) application.
        ''' </summary>
        ConsoleApplication

        ''' <summary>
        ''' The application is a Windows-based application.
        ''' </summary>
        WindowApplication

    End Enum

End Namespace

#End If

#End Region
