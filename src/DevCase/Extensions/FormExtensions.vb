' This source-code is freely distributed as part of "DevCase for .NET Framework".
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

#End Region

#Region " Form Extensions "

Namespace DevCase.Core.Extensions

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains custom extension methods to use with the <see cref="Form"/> type.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    <HideModuleName>
    Friend Module FormExtensions

#Region " Public Extension Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Changes the color appearance of the source <see cref="Form"/> to its default appearance.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="f">
        ''' The source <see cref="Form"/>.
        ''' </param>
        ''' 
        ''' <param name="childControls">
        ''' <see langword="True"/> to change the color appearance of child controls too.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        <Extension>
        <EditorBrowsable(EditorBrowsableState.Always)>
        Public Sub SetThemeDefault(ByVal f As Form, ByVal childControls As Boolean)
            f.BackColor = Form.DefaultBackColor
            f.ForeColor = Form.DefaultForeColor

            If (childControls) Then
                ForEachControl(f, True, Sub(ctrl As Control)
                                            If (ctrl.GetType().IsPublic) Then
                                                ctrl.SetThemeDefault()
                                            End If
                                        End Sub)
            End If
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Changes the color appearance of the source <see cref="Form"/> to Visual Studio Dark Theme appearance.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="f">
        ''' The source <see cref="Form"/>.
        ''' </param>
        ''' 
        ''' <param name="childControls">
        ''' <see langword="True"/> to change the color appearance of child controls too.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        <Extension>
        <EditorBrowsable(EditorBrowsableState.Always)>
        Public Sub SetThemeVisualStudioDark(ByVal f As Form, ByVal childControls As Boolean)
            f.BackColor = Color.FromArgb(255, 45, 45, 48)
            f.ForeColor = Color.Gainsboro

            If (childControls) Then
                ForEachControl(f, True, Sub(ctrl As Control)
                                            If (ctrl.GetType().IsPublic) Then
                                                ctrl.SetThemeVisualStudioDark()
                                            End If
                                        End Sub)
            End If
        End Sub

#End Region

    End Module

End Namespace

#End Region
