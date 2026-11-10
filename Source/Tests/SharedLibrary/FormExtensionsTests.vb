#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.Drawing
Imports System.Windows.Forms

Imports DevCase.Core.Application.UserInterface
Imports DevCase.Core.Extensions

Imports Microsoft.VisualStudio.TestTools.UnitTesting

#End Region

#Region " FormExtensionsTests "

Namespace EasyLink.Tests.SharedLibrary

    ''' <summary>
    ''' Unit tests for the <see cref="FormExtensions"/> visual theme extensions.
    ''' <para></para>
    ''' The forms are never shown on screen: only their window handles are created when needed.
    ''' </summary>
    <STATestClass>
    Public NotInheritable Class FormExtensionsTests

#Region " SetVisualTheme "

        ''' <summary>
        ''' Verifies that applying a theme to a form whose handle does not exist yet sets the colors
        ''' but does not force the creation of the window handle (the title bar mode is applied later,
        ''' when the handle is created).
        ''' </summary>
        <TestMethod>
        Public Sub SetVisualTheme_HandleNotCreated_DoesNotCreateHandle()

            Using f As New Form()
                f.SetVisualTheme(VisualTheme.VisualStudioDark, childControls:=True)

                Assert.IsFalse(f.IsHandleCreated)
                Assert.AreEqual(Color.FromArgb(255, 45, 45, 48), f.BackColor)
                Assert.AreEqual(Color.Gainsboro, f.ForeColor)
            End Using

        End Sub

        ''' <summary>
        ''' Verifies that the default theme restores the default form colors.
        ''' </summary>
        <TestMethod>
        Public Sub SetVisualTheme_Default_RestoresDefaultColors()

            Using f As New Form()
                f.SetVisualTheme(VisualTheme.VisualStudioDark, childControls:=False)
                f.SetVisualTheme(VisualTheme.Default, childControls:=False)

                Assert.AreEqual(Form.DefaultBackColor, f.BackColor)
                Assert.AreEqual(Form.DefaultForeColor, f.ForeColor)
            End Using

        End Sub

        ''' <summary>
        ''' Verifies that the title bar mode is re-applied without errors every time the window handle is
        ''' created or recreated (as happens with forms shown with ShowDialog), after the theme was set
        ''' before the handle existed.
        ''' </summary>
        <TestMethod>
        Public Sub SetVisualTheme_HandleCreatedAndRecreated_ReappliesWithoutErrors()

            Using f As New Form()
                f.SetVisualTheme(VisualTheme.VisualStudioDark, childControls:=False)

                Dim firstHandle As IntPtr = f.Handle ' Creates the handle (raises HandleCreated).
                Assert.AreNotEqual(IntPtr.Zero, firstHandle)

                f.ShowInTaskbar = Not f.ShowInTaskbar ' Forces the handle to be recreated.
                Dim secondHandle As IntPtr = f.Handle

                Assert.IsTrue(f.IsHandleCreated)
                Assert.AreNotEqual(firstHandle, secondHandle)
            End Using

        End Sub

        ''' <summary>
        ''' Verifies that changing the theme of a form whose handle exists but that is not on screen
        ''' does not resize it (the 1-pixel resize workaround is only needed for windows already on screen,
        ''' and applying it before the window is shown causes needless layout passes and Resize events).
        ''' </summary>
        <TestMethod>
        Public Sub SetVisualTheme_HandleCreatedNotOnScreen_DoesNotResize()

            Using f As New Form() With {.Size = New Size(640, 480)}
                Dim handle As IntPtr = f.Handle
                Dim resizeCount As Integer = 0
                AddHandler f.Resize, Sub(sender As Object, e As EventArgs) resizeCount += 1

                f.SetVisualTheme(VisualTheme.VisualStudioDark, childControls:=False)
                f.SetVisualTheme(VisualTheme.Default, childControls:=False)

                Assert.AreEqual(handle, f.Handle)
                Assert.AreEqual(New Size(640, 480), f.Size)
                Assert.AreEqual(0, resizeCount)
            End Using

        End Sub

        ''' <summary>
        ''' Verifies that a <see langword="Nothing"/> form is rejected.
        ''' </summary>
        <TestMethod>
        Public Sub SetVisualTheme_NullForm_ThrowsArgumentNullException()

            Dim f As Form = Nothing

            Assert.ThrowsExactly(Of ArgumentNullException)(
                Sub() FormExtensions.SetVisualTheme(f, VisualTheme.Default, childControls:=False))

        End Sub

#End Region

    End Class

End Namespace

#End Region
