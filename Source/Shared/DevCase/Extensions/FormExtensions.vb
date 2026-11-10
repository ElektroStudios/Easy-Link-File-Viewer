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
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Imports DevCase.Core.Application.UserInterface
Imports DevCase.Extensions
Imports DevCase.Win32
Imports DevCase.Win32.Enums

#End Region

#Region " Form Extensions "

Namespace DevCase.Core.Extensions

    ''' 
    ''' <summary>
    ''' Contains custom extension methods to use with the <see cref="Form"/> type.
    ''' </summary>
    ''' 
    <HideModuleName>
    Public Module FormExtensions

#Region " Public Extension Methods "

        ''' <summary>
        ''' Iterates through all controls within a parent <see cref="Form"/>, 
        ''' optionally recursively, and performs the specified action on each control.
        ''' </summary>
        '''
        ''' <param name="form">
        ''' The parent <see cref="Form"/> whose child controls are to be iterated.
        ''' </param>
        ''' 
        ''' <param name="recursive">
        ''' <see langword="True"/> to iterate recursively through all child controls 
        ''' (i.e., iterate the child controls of child controls); otherwise, <see langword="False"/>.
        ''' </param>
        ''' 
        ''' <param name="action">
        ''' The action to perform on each control.
        ''' </param>
        <DebuggerStepThrough>
        <Extension>
        <EditorBrowsable(EditorBrowsableState.Always)>
        Public Sub ForEachControl(form As Form, recursive As Boolean, action As Action(Of Control))

            ContainerControlExtensions.ForEachControl(Of Control)(form, recursive, action)

        End Sub

        ''' <summary>
        ''' Changes the color appearance of the source <see cref="Form"/> using the specified theme.
        ''' <para></para>
        ''' The dark/light title bar (DWM immersive dark mode attribute) is a per-window-handle setting.
        ''' This method remembers the theme assigned to the form, and automatically re-applies the
        ''' title bar mode every time the form handle is (re)created. This is required, for example,
        ''' for forms shown with <see cref="Form.ShowDialog()"/>, whose handle is destroyed when the
        ''' dialog is closed and created again the next time it is shown.
        ''' </summary>
        '''
        ''' <param name="f">
        ''' The source <see cref="Form"/>.
        ''' </param>
        ''' 
        ''' <param name="theme">
        ''' The visual theme.
        ''' </param>
        ''' 
        ''' <param name="childControls">
        ''' <see langword="True"/> to change the color appearance of child controls too.
        ''' </param>
        <DebuggerStepThrough>
        <Extension>
        <EditorBrowsable(EditorBrowsableState.Always)>
        Public Sub SetVisualTheme(f As Form, theme As VisualTheme, childControls As Boolean)

            If f Is Nothing Then
                Throw New ArgumentNullException(NameOf(f))
            End If

            Dim useDarkMode As Boolean

            Select Case theme

                Case VisualTheme.Default
                    f.BackColor = Form.DefaultBackColor
                    f.ForeColor = Form.DefaultForeColor
                    useDarkMode = False

                Case VisualTheme.VisualStudioDark
                    f.BackColor = System.Drawing.Color.FromArgb(255, 45, 45, 48)
                    f.ForeColor = System.Drawing.Color.Gainsboro
                    useDarkMode = True

                Case Else
                    Throw New InvalidEnumArgumentException(NameOf(theme), theme, GetType(VisualTheme))

            End Select

            If childControls Then
                FormExtensions.ForEachControl(f, True,
                                              Sub(ctrl As Control)
                                                  If ctrl.GetType().IsPublic Then
                                                      ControlExtensions.SetVisualTheme(ctrl, theme)
                                                  End If
                                              End Sub)
            End If

            ' Remember the title bar mode of this form, and subscribe (only once per form)
            ' to its HandleCreated event to re-apply it whenever the handle is (re)created.
            Dim darkModeState As StrongBox(Of Boolean) = Nothing
            If FormExtensions.titleBarDarkModeStates.TryGetValue(f, darkModeState) Then
                darkModeState.Value = useDarkMode
            Else
                FormExtensions.titleBarDarkModeStates.Add(f, New StrongBox(Of Boolean)(useDarkMode))
                AddHandler f.HandleCreated, AddressOf FormExtensions.Form_HandleCreated
            End If

            ' Do not force the creation of the window handle here (accessing the Form.Handle property would do it).
            ' If the handle does not exist yet, the title bar mode will be applied when it is created.
            If f.IsHandleCreated Then
                FormExtensions.ApplyTitleBarDarkMode(f, useDarkMode)
            End If

        End Sub

#End Region

#Region " Private Fields "

        ''' <summary>
        ''' Keeps the title bar dark mode state assigned to each form by
        ''' <see cref="FormExtensions.SetVisualTheme(Form, VisualTheme, Boolean)"/>,
        ''' without preventing the forms from being garbage collected.
        ''' </summary>
        Private ReadOnly titleBarDarkModeStates As New ConditionalWeakTable(Of Form, StrongBox(Of Boolean))

#End Region

#Region " Private Methods "

        ''' <summary>
        ''' Handles the <see cref="Control.HandleCreated"/> event of a form whose visual theme was set by
        ''' <see cref="FormExtensions.SetVisualTheme(Form, VisualTheme, Boolean)"/>,
        ''' re-applying the title bar dark mode to the new window handle.
        ''' </summary>
        ''' 
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        ''' 
        ''' <param name="e">
        ''' The <see cref="EventArgs"/> instance containing the event data.
        ''' </param>
        <DebuggerStepThrough>
        Private Sub Form_HandleCreated(sender As Object, e As EventArgs)

            Dim f As Form = DirectCast(sender, Form)
            Dim darkModeState As StrongBox(Of Boolean) = Nothing
            If Not FormExtensions.titleBarDarkModeStates.TryGetValue(f, darkModeState) Then
                Throw New InvalidOperationException($"No visual theme state was registered for form '{f.Name}'.")
            End If

            FormExtensions.ApplyTitleBarDarkMode(f, darkModeState.Value)

        End Sub

        ''' <summary>
        ''' Applies the dark or light title bar (DWM immersive dark mode attribute) to the window handle of the specified form,
        ''' and forces the non-client area to be redrawn so the change takes effect.
        ''' </summary>
        ''' 
        ''' <param name="f">
        ''' The source <see cref="Form"/>. Its handle must be already created.
        ''' </param>
        ''' 
        ''' <param name="useDarkMode">
        ''' <see langword="True"/> to use a dark title bar; <see langword="False"/> to use a light title bar.
        ''' </param>
        <DebuggerStepThrough>
        Private Sub ApplyTitleBarDarkMode(f As Form, useDarkMode As Boolean)

            Const DWMWA_USE_IMMERSIVE_DARK_MODE As Integer = 20
            Dim useDarkModeValue As Integer = If(useDarkMode, 1, 0)
            Dim useDarkModeResult As HResult =
                NativeMethods.DwmSetWindowAttribute(f.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE,
                                                    useDarkModeValue, Marshal.SizeOf(useDarkModeValue))

            ' The attribute is not supported on Windows versions prior to Windows 10 20H1 (build 19041).
            ' In that case the title bar simply keeps the system appearance, as before.
            If useDarkModeResult <> HResult.S_OK Then
                Exit Sub
            End If

            ' Note: Form.Visible cannot be used here, because Windows Forms raises the Load event of top-level forms
            ' while processing WM_SHOWWINDOW, when Form.Visible already returns True but the window is not on screen yet.
            If Not NativeMethods.IsWindowVisible(f.Handle) Then
                ' The window is not on screen yet (e.g. it is being created or loaded):
                ' notify a frame change so the new caption mode is used when the window is shown.
                ' Unlike the resize workaround below, this does not change the window size,
                ' so it does not trigger extra layout passes, Resize events or repaints.
                Const flags As SetWindowPosFlags = SetWindowPosFlags.FrameChanged Or
                                                   SetWindowPosFlags.IgnoreMove Or
                                                   SetWindowPosFlags.IgnoreResize Or
                                                   SetWindowPosFlags.IgnoreZOrder Or
                                                   SetWindowPosFlags.DontActivate
                If Not NativeMethods.SetWindowPos(f.Handle, IntPtr.Zero, 0, 0, 0, 0, flags) Then
                    Throw New Win32Exception(Marshal.GetLastWin32Error())
                End If
                Exit Sub
            End If

            ' The window is already on screen (e.g. the theme was changed at runtime).
            ' If the window is in normal state, force it to redraw by
            ' resizing it by 1 pixel and then resizing it back to the original size.
            ' This is required because the caption color change (DwmWindowAttribute.UseImmersiveDarkMode)
            ' is not repainted until the window frame is resized.
            If f.WindowState = FormWindowState.Normal Then
                Dim originalSize As Size = f.Size

                ' --- Determine a safe delta (+1 or -1) ---
                Dim delta As Integer = 0

                ' Check if we can grow by 1 pixel in height
                Dim canGrow As Boolean =
                    (f.MaximumSize.Height = 0 OrElse originalSize.Height + 1 <= f.MaximumSize.Height) AndAlso
                    (f.MaximumSize.Width = 0 OrElse originalSize.Width <= f.MaximumSize.Width)

                ' Check if we can shrink by 1 pixel in height
                Dim canShrink As Boolean =
                    (originalSize.Height - 1 >= f.MinimumSize.Height) AndAlso
                    (originalSize.Height - 1 >= SystemInformation.MinimumWindowSize.Height)

                If canGrow Then
                    delta = 1
                ElseIf canShrink Then
                    delta = -1
                End If

                ' Only attempt to resize if we found a delta.
                If delta <> 0 Then
                    f.Size = New Size(originalSize.Width, originalSize.Height + delta)
                    f.Size = originalSize
                End If
            End If

        End Sub

#End Region

    End Module

End Namespace

#End Region
