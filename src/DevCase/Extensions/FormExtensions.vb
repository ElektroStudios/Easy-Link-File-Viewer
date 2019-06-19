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
Imports DevCase.Core.Application.UserInterface

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
        ''' Changes the color appearance of the source <see cref="Form"/> using the specified style.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="f">
        ''' The source <see cref="Form"/>.
        ''' </param>
        ''' 
        ''' <param name="style">
        ''' The visual style.
        ''' </param>
        ''' 
        ''' <param name="childControls">
        ''' <see langword="True"/> to change the color appearance of child controls too.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        <Extension>
        <EditorBrowsable(EditorBrowsableState.Always)>
        Public Sub SetVisualStyle(ByVal f As Form, ByVal style As VisualStyle, ByVal childControls As Boolean)

            Select Case style

                Case VisualStyle.Default
                    f.BackColor = Form.DefaultBackColor
                    f.ForeColor = Form.DefaultForeColor


                Case VisualStyle.VisualStudioDark
                    f.BackColor = Color.FromArgb(255, 45, 45, 48)
                    f.ForeColor = Color.Gainsboro

                Case Else
                    Throw New InvalidEnumArgumentException(NameOf(style), style, GetType(VisualStyle))

            End Select

            If (childControls) Then
                f.ForEachControl(True,
                                              Sub(ctrl As Control)
                                                  If (ctrl.GetType().IsPublic) Then
                                                      ctrl.SetVisualStyle(style)
                                                  End If
                                              End Sub)
            End If
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Iterate through all the controls in the source <see cref="Form"/> 
        ''' and performs the specified action on each one.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="f">
        ''' The source <see cref="Form"/>.
        ''' </param>
        ''' 
        ''' <param name="recursive">
        ''' If <see langword="True"/>, iterates through controls of child controls too.
        ''' </param>
        ''' 
        ''' <param name="action">
        ''' An <see cref="Action(Of Control)"/> to perform on each control.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        <Extension>
        <EditorBrowsable(EditorBrowsableState.Always)>
        Public Sub ForEachControl(ByVal f As Form, ByVal recursive As Boolean, ByVal action As Action(Of Control))
            f.ForEachControl(Of Control)(recursive, action)
        End Sub

#End Region

    End Module

End Namespace

#End Region
