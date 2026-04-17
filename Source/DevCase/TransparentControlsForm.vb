' This source-code is freely distributed as part of "DevCase Class Library .NET Developers".
' 
' Maybe you would like to consider to buy this powerful set of libraries to support me. 
' You can do loads of things with my apis for a big amount of diverse thematics.
' 
' Here is a link to the purchase page:
' https://codecanyon.net/item/elektrokit-class-library-for-net/19260282
' 
' Thank you.

#Region " Usage Examples "

'Public Class Form1 : Inherits Form
'
'' Our Transparent Form.
'Protected WithEvents frm As TransparentControlsForm
'
'' Our controls.
'Friend WithEvents tb1 As New TextBox With {.Text = "Elektro-Test 1"}
'Friend WithEvents tb2 As New TextBox With {.Text = "Elektro-Test 2"}
'Friend WithEvents cb1 As New CheckBox With {.Text = "Elektro-Test 3", .FlatStyle = FlatStyle.Flat}
'
'#Region " Event Handlers "
'
'    ''' <summary>
'    ''' Handles the Shown event of the Form.
'    ''' </summary>
'    Private Shadows Sub Shown(sender As Object, e As EventArgs) _
'    Handles MyBase.Shown
'
'' Set the Control locations.
'        tb1.Location = New Point(5, 5)
'        tb2.Location = New Point(tb1.Location.X, tb1.Location.Y + CInt(tb1.Height * 1.5R))
'        cb1.Location = New Point(tb2.Location.X, tb2.Location.Y + CInt(tb2.Height * 1.5R))
'
'' Instance the Form that will store our controls.
'        If frm Is Nothing Then
'            frm = New TransparentControlsForm({tb1, tb2, cb1},
'                       New Point(Me.Bounds.Right, Me.Bounds.Top))
'        End If
'
'       With frm
'           .Moveable = True ' Set the Controls moveable.
'           .Show() ' Display the transparent Form.
'       End With
'
'    End Sub
'
'#End Region
'
'#Region " Textbox's Event Handlers "
'
'    ''' <summary>
'    ''' Handles the TextChanged event of the TextBox controls.
'    ''' </summary>
'    ''' <param name="sender">The source of the event.</param>
'    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
'    Friend Sub tb_textchanged(sender As Object, e As KeyPressEventArgs) _
'    Handles tb1.KeyPress, tb2.KeyPress, cb1.KeyPress
'
'        ' Just a crazy message-box to interacts with the raised event.
'        MessageBox.Show(I'm gonna do this control disappear!", "",
'                        MessageBoxButtons.OK, MessageBoxIcon.Stop)
'
'        e.Handled = True
'
'        ' Searchs the ControlsForm
'        Dim f As Form = DirectCast(sender, Control).FindForm
'
'        ' ...And close it
'        f.Hide()
'
'    End Sub
'
'#End Region
'
'End Class

#End Region

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.ComponentModel

#End Region

#Region " TransparentControls Form "

' ReSharper disable once CheckNamespace

Namespace DevCase.Core.Application.Forms

    ''' <summary>
    ''' A transparent <see cref="Form"/> designed to host controls.
    ''' </summary>
    '''
    ''' <example> This is a code example.
    ''' <code language="VB.NET">
    ''' Public Class Form1 : Inherits Form
    ''' 
    '''     ' Our Transparent Form.
    '''     Protected WithEvents frm As TransparentControlsForm
    ''' 
    '''     ' Our controls.
    '''     Friend WithEvents tb1 As New TextBox With {.Text = "Elektro-Test 1"}
    '''     Friend WithEvents tb2 As New TextBox With {.Text = "Elektro-Test 2"}
    '''     Friend WithEvents cb1 As New CheckBox With {.Text = "Elektro-Test 3", .FlatStyle = FlatStyle.Flat}
    ''' 
    ''' #Region " Event Handlers "
    ''' 
    '''     ''' &lt;summary&gt;
    '''     ''' Handles the Shown event of the Form.
    '''     ''' &lt;/summary&gt;
    '''     Private Shadows Sub Shown(sender As Object, e As EventArgs) _
    '''     Handles MyBase.Shown
    ''' 
    '''         ' Set the Control locations.
    '''         tb1.Location = New Point(5, 5)
    '''         tb2.Location = New Point(tb1.Location.X, tb1.Location.Y + CInt(tb1.Height * 1.5R))
    '''         cb1.Location = New Point(tb2.Location.X, tb2.Location.Y + CInt(tb2.Height * 1.5R))
    ''' 
    '''         ' Instance the Form that will store our controls.
    '''         If frm Is Nothing Then
    '''             frm = New TransparentControlsForm({tb1, tb2, cb1},
    '''                        New Point(Me.Bounds.Right, Me.Bounds.Top))
    '''         End If
    ''' 
    '''         With frm
    '''             .Moveable = True ' Set the Controls moveable.
    '''             .Show() ' Display the transparent Form.
    '''         End With
    ''' 
    '''     End Sub
    ''' 
    ''' #End Region
    ''' 
    ''' #Region " Textbox's Event Handlers "
    ''' 
    '''     ''' &lt;summary&gt;
    '''     ''' Handles the TextChanged event of the TextBox controls.
    '''     ''' &lt;/summary&gt;
    '''     ''' &lt;param name="sender"&gt;The source of the event.&lt;/param&gt;
    '''     ''' &lt;param name="e"&gt;The <see cref="EventArgs"/> instance containing the event data.&lt;/param&gt;
    '''     Friend Sub tb_textchanged(sender As Object, e As KeyPressEventArgs) _
    '''     Handles tb1.KeyPress, tb2.KeyPress, cb1.KeyPress
    ''' 
    '''         ' Just a crazy message-box to interact with the raised event.
    '''         MessageBox.Show("I'm gonna do this control disappear!", "",
    '''                         MessageBoxButtons.OK, MessageBoxIcon.Stop)
    ''' 
    '''         e.Handled = True
    ''' 
    '''         ' Searchs the ControlsForm
    '''         Dim f As Form = DirectCast(sender, Control).FindForm
    ''' 
    '''         ' ...And close it
    '''         f.Hide()
    ''' 
    '''     End Sub
    ''' 
    ''' #End Region
    ''' 
    ''' End Class
    ''' </code>
    ''' </example>
    <ToolboxItem(False)>
    <DesignTimeVisible(False)>
    <DesignerCategory(NameOf(DesignerCategoryAttribute.Generic))>
    Public Class TransparentControlsForm : Inherits Form

#Region " Private Fields "

        ''' <summary>
        ''' Indicates whether the moveable events are handled
        ''' </summary>
        Protected moveableIsHandled As Boolean

        ''' <summary>
        ''' Boolean Flag that indicates whether the Form should be moved.
        ''' </summary>
        Protected moveFormFlag As Boolean

        ''' <summary>
        ''' The position where to move the form.
        ''' </summary>
        Protected moveFormPosition As Point

#End Region

#Region " Properties "

        ''' <summary>
        ''' Gets or sets a value indicating whether this <see cref="TransparentControlsForm"/> and it's controls are movable.
        ''' </summary>
        '''
        ''' <value>
        ''' <see langword="True"/> if controls are movable; otherwise, <see langword="False"/>.
        ''' </value>
        Public Overridable Property Moveable As Boolean

            <DebuggerStepThrough>
            Get
                Return Me.moveableB
            End Get

            <DebuggerStepThrough>
            Set(value As Boolean)

                Me.moveableB = value

                Dim pan As Panel =
                  (From p As Panel In MyBase.Controls.OfType(Of Panel)()
                   Where DirectCast(p.Tag, IntPtr) = Me.Handle).First

                Select Case value

                    Case True ' Add Moveable Events to EventHandler.

                        If Not Me.moveableIsHandled Then ' If not Moveable handlers are already handled then...
                            For Each c As Control In pan.Controls
                                AddHandler c.MouseDown, AddressOf Me.MouseDown
                                AddHandler c.MouseUp, AddressOf Me.MouseUp
                                AddHandler c.MouseMove, AddressOf Me.MouseMove
                            Next c
                            Me.moveableIsHandled = True
                        End If

                    Case False ' Remove Moveable Events from EventHandler.

                        If Me.moveableIsHandled Then ' If Moveable handlers are already handled then...
                            For Each c As Control In pan.Controls
                                RemoveHandler c.MouseDown, AddressOf Me.MouseDown
                                RemoveHandler c.MouseUp, AddressOf Me.MouseUp
                                RemoveHandler c.MouseMove, AddressOf Me.MouseMove
                            Next c
                            Me.moveableIsHandled = False
                        End If

                End Select

            End Set

        End Property

        ''' <summary>
        ''' ( Backing field )
        ''' A value indicating whether this <see cref="TransparentControlsForm"/> and it's controls are movable.
        ''' </summary>
        Protected moveableB As Boolean = False

#End Region

#Region " Constructors "

        ''' <summary>
        ''' Prevents a default instance of the <see cref="TransparentControlsForm"/> class from being created.
        ''' </summary>
        <DebuggerNonUserCode>
        Private Sub New()
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the <see cref="TransparentControlsForm"/> class.
        ''' </summary>
        '''
        ''' <param name="controls">
        ''' The control array to display in the Form.
        ''' </param>
        ''' 
        ''' <param name="formLocation">
        ''' The default Formulary location.
        ''' </param>
        <DebuggerStepThrough>
        Public Sub New(controls As Control(),
                       Optional formLocation As Point = Nothing)

            ' InitializeComponent call.
            MyBase.SuspendLayout()
            MyBase.Name = "ControlsForm"

            ' Adjust Form size.
            MyBase.ClientSize = New Size(0, 0)
            MyBase.AutoSize = True
            MyBase.AutoSizeMode = AutoSizeMode.GrowAndShrink

            ' Set the Transparency properties.
            MyBase.AllowTransparency = True
            MyBase.BackColor = Color.Fuchsia
            MyBase.TransparencyKey = Color.Fuchsia
            MyBase.DoubleBuffered = False

            ' Is not necessary to display borders, icon, and taskbar, hide them.
            MyBase.FormBorderStyle = FormBorderStyle.None
            MyBase.ShowIcon = False
            MyBase.ShowInTaskbar = False

            ' Instance a Panel to add our controls.
            Dim pan As New Panel
            With pan

                ' Suspend the Panel layout.
                pan.SuspendLayout()

                ' Set the Panel properties.
                .Name = "ControlsForm Panel"
                .Tag = Me.Handle
                .AutoSize = True
                .AutoSizeMode = AutoSizeMode.GrowAndShrink
                .BorderStyle = BorderStyle.None
                .Dock = DockStyle.Fill

            End With

            ' Add our controls inside the Panel.
            pan.Controls.AddRange(controls)

            ' Add the Panel inside the Form.
            MyBase.Controls.Add(pan)

            ' If custom Form location is set then...
            If Not formLocation.Equals(Nothing) Then

                ' Set the StartPosition to manual relocation.
                MyBase.StartPosition = FormStartPosition.Manual

                ' Relocate the form.
                MyBase.Location = formLocation

            End If

            ' Resume layouts.
            pan.ResumeLayout(False)
            MyBase.ResumeLayout(False)

        End Sub

        ''' <summary>
        ''' Initializes a new instance of the <see cref="TransparentControlsForm"/> class.
        ''' </summary>
        '''
        ''' <param name="control">
        ''' The control to display in the Formulary.
        ''' </param>
        ''' 
        ''' <param name="formLocation">
        ''' The default Formulary location.
        ''' </param>
        <DebuggerStepThrough>
        Public Sub New(control As Control,
                       Optional formLocation As Point = Nothing)

            Me.New({control}, formLocation)

        End Sub

#End Region

#Region " Event Handlers "

        ''' <summary>
        ''' Handles the <see cref="TransparentControlsForm.MouseDown"/> event of the Form and it's controls.
        ''' </summary>
        '''
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        ''' 
        ''' <param name="e">
        ''' The <see cref="MouseEventArgs"/> instance containing the event data.
        ''' </param>
        <DebuggerStepThrough>
        Protected Shadows Sub MouseDown(sender As Object, e As MouseEventArgs) _
        Handles MyBase.MouseDown

            If e.Button = MouseButtons.Left Then
                Me.moveFormFlag = True
                Me.Cursor = Cursors.NoMove2D
                Me.moveFormPosition = e.Location
            End If

        End Sub

        ''' <summary>
        ''' Handles the <see cref="TransparentControlsForm.MouseMove"/> event of the Form and it's controls.
        ''' </summary>
        '''
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        ''' 
        ''' <param name="e">
        ''' The <see cref="MouseEventArgs"/> instance containing the event data.
        ''' </param>
        <DebuggerStepThrough>
        Protected Shadows Sub MouseMove(sender As Object, e As MouseEventArgs) _
        Handles MyBase.MouseMove

            If Me.moveFormFlag Then
                Me.Location += New Size(e.Location.X - Me.moveFormPosition.X, e.Location.Y - Me.moveFormPosition.Y)
            End If

        End Sub

        ''' <summary>
        ''' Handles the <see cref="TransparentControlsForm.MouseUp"/> event of the Form and it's controls.
        ''' </summary>
        '''
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        ''' 
        ''' <param name="e">
        ''' The <see cref="MouseEventArgs"/> instance containing the event data.
        ''' </param>
        <DebuggerStepThrough>
        Protected Shadows Sub MouseUp(sender As Object, e As MouseEventArgs) _
        Handles MyBase.MouseUp

            If e.Button = MouseButtons.Left Then
                Me.moveFormFlag = False
                Me.Cursor = Cursors.Default
            End If

        End Sub

#End Region

    End Class

End Namespace

#End Region
