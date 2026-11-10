#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.IO

#End Region

#Region " ConsoleCapture "

Namespace EasyLink.Tests.Helpers

    ''' <summary>
    ''' Redirects the standard output and standard error streams of the <see cref="Console"/>
    ''' to in-memory writers, and restores the original streams when disposed.
    ''' </summary>
    Friend NotInheritable Class ConsoleCapture : Implements IDisposable

#Region " Private Fields "

        ''' <summary>
        ''' The original standard output writer.
        ''' </summary>
        Private ReadOnly originalOut As TextWriter

        ''' <summary>
        ''' The original standard error writer.
        ''' </summary>
        Private ReadOnly originalError As TextWriter

        ''' <summary>
        ''' The writer that receives the redirected standard output.
        ''' </summary>
        Private ReadOnly outWriter As StringWriter

        ''' <summary>
        ''' The writer that receives the redirected standard error.
        ''' </summary>
        Private ReadOnly errorWriter As StringWriter

        ''' <summary>
        ''' Flag to detect redundant calls to <see cref="ConsoleCapture.Dispose()"/>.
        ''' </summary>
        Private isDisposed As Boolean

#End Region

#Region " Properties "

        ''' <summary>
        ''' Gets the text written to the standard output stream so far.
        ''' </summary>
        Friend ReadOnly Property Output As String
            Get
                Return Me.outWriter.ToString()
            End Get
        End Property

        ''' <summary>
        ''' Gets the text written to the standard error stream so far.
        ''' </summary>
        Friend ReadOnly Property ErrorOutput As String
            Get
                Return Me.errorWriter.ToString()
            End Get
        End Property

#End Region

#Region " Constructors "

        ''' <summary>
        ''' Initializes a new instance of the <see cref="ConsoleCapture"/> class
        ''' and starts redirecting the console streams.
        ''' </summary>
        Friend Sub New()

            Me.originalOut = Console.Out
            Me.originalError = Console.Error

            Me.outWriter = New StringWriter()
            Me.errorWriter = New StringWriter()

            Console.SetOut(Me.outWriter)
            Console.SetError(Me.errorWriter)

        End Sub

#End Region

#Region " IDisposable Implementation "

        ''' <summary>
        ''' Restores the original console streams and releases the in-memory writers.
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose

            If Me.isDisposed Then
                Exit Sub
            End If

            Console.SetOut(Me.originalOut)
            Console.SetError(Me.originalError)

            Me.outWriter.Dispose()
            Me.errorWriter.Dispose()

            Me.isDisposed = True

        End Sub

#End Region

    End Class

End Namespace

#End Region
