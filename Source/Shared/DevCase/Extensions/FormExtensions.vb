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

            Dim useDarkMode As Integer

            Select Case theme

                Case VisualTheme.Default
                    f.BackColor = Form.DefaultBackColor
                    f.ForeColor = Form.DefaultForeColor
                    useDarkMode = 0

                Case VisualTheme.VisualStudioDark
                    f.BackColor = System.Drawing.Color.FromArgb(255, 45, 45, 48)
                    f.ForeColor = System.Drawing.Color.Gainsboro
                    useDarkMode = 1

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

            Const DWMWA_USE_IMMERSIVE_DARK_MODE As Integer = 20
            Dim useDarkModeResult As HResult =
                NativeMethods.DwmSetWindowAttribute(f.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE,
                                                    useDarkMode, Marshal.SizeOf(useDarkMode))

            ' If the window is in normal state, force it to redraw by
            ' resizing it by 1 pixel and then resizing it back to the original size.
            ' This is required because the caption color change (DwmWindowAttribute.UseImmersiveDarkMode).
            If (useDarkModeResult = HResult.S_OK) AndAlso (f.WindowState = FormWindowState.Normal) Then
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
