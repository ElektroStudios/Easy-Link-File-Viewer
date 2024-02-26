' ***********************************************************************
' Author   : ElektroStudios
' Modified : 29-April-2017
' ***********************************************************************

#Region " Public Members Summary "

#Region " Methods "

' OpenInExplorer(String)
' OpenInExplorer(FileInfo)

#End Region

#End Region

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.IO

#End Region

#Region " Files "

Namespace DevCase.Core.IO.Tools

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains file related utilities.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Friend NotInheritable Class FileUtil

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Prevents a default instance of the <see cref="FileUtil"/> class from being created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub New()

        End Sub

#End Region

#Region " Public Methods "

#End Region

#Region " Private Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Opens the specified file or folder in Explorer.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="itemPath">
        ''' The file or folder path.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="FileNotFoundException">
        ''' File or directory not found.;itemPath
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Friend Shared Sub InternalOpenInExplorer(itemPath As String)

            Dim arguments As String

            If File.Exists(itemPath) Then
                Dim file As New FileInfo(itemPath)
                arguments = String.Format("/Select,""{0}""", Path.Combine(file.DirectoryName, file.Name))

            ElseIf Directory.Exists(itemPath) Then
                Dim dir As New DirectoryInfo(itemPath)
                arguments = String.Format("""{0}""", dir.FullName)

            Else
                Throw New FileNotFoundException("File or directory not found.", fileName:=itemPath)

            End If

            Using p As Process = Process.Start("Explorer.exe", arguments)
            End Using

        End Sub

#End Region

    End Class

End Namespace

#End Region
