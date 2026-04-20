#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

Public NotInheritable Class Program

    Private Sub New()
    End Sub

    <STAThread>
    Public Shared Sub Main(args As String())

        Dim processor As New CommandLineProcessor()
        Dim exitCode As Integer = processor.Run(args)
        Environment.ExitCode = exitCode
    End Sub

End Class
