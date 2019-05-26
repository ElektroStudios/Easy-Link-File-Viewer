Imports System.IO
Imports DevCase.Core.IO

Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim lnk As New ShortcutFileInfo("C:\Test Shortcut.lnk")
        lnk.Create()

        lnk.IsReadOnly = False ' Deletes FileAttributes.ReadOnly attribute from file.
        With lnk
            .Attributes = FileAttributes.System
            .Description = "My shortcut file description"
            .Hotkey = Keys.Shift Or Keys.Alt Or Keys.Control Or Keys.F1
            .Icon = "Shell32.dll"
            .IconIndex = 0
            .Target = "C:\Target.exe"
            .TargetArguments = "/arg1 /arg2"
            .WindowState = ShortcutWindowState.Normal
            .WorkingDirectory = Path.GetDirectoryName(lnk.Target)
        End With
        lnk.IsReadOnly = True ' Adds FileAttributes.ReadOnly attribute to file.

        Me.PropertyGrid1.SelectedObject = lnk

    End Sub

    Private Sub PropertyGrid1_PropertyValueChanged(s As Object, e As PropertyValueChangedEventArgs) Handles PropertyGrid1.PropertyValueChanged

    End Sub

End Class
