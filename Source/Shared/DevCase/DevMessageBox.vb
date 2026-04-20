' ***********************************************************************
' Author   : ElektroStudios
' Modified : 18-April-2026
' ***********************************************************************

#Region " Usage Examples "

#Region " Static Methods "


'' This is a code example that demonstrates static methods usage from a Windows Forms application.
'
' Dim result As DialogResult = DevMessageBox.ShowStatic(Me, "Text", "Caption", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
' Console.WriteLine(result.ToString())


'' This is a code example that demonstrates static methods usage from a Console application.
'
' Dim ownerWindow As New ConsoleWindowWrapper(UtilConsole.Handle)
' DevCase.Win32.NativeMethods.SetForegroundWindow(ownerWindow.Handle)
'
' Dim result As DialogResult = DevMessageBox.ShowStatic(ownerWindow, "Text", "Caption", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
' Console.WriteLine(result.ToString())


#End Region

#Region " Instance Methods "


'' This is a code example that demonstrates instance methods usage from a Windows Forms application.
'
'Using msg As New DevMessageBox(owner:=Me, TimeSpan.FromSeconds(5))
'
'    Dim dlgresult As DialogResult = msg.Show("Question", "Title", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
'    Console.WriteLine(dlgresult.ToString())
'End Using


'' This is a code example that demonstrates instance methods usage from a Console application.
'
' Dim ownerWindow As New ConsoleWindowWrapper(UtilConsole.Handle)
' DevCase.Win32.NativeMethods.SetForegroundWindow(ownerWindow.Handle)
'
'Using msg As New DevMessageBox(ownerWindow, TimeSpan.FromSeconds(5))
'
'    Dim dlgresult As DialogResult = msg.Show("Question", "Title", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
'    Console.WriteLine(dlgresult.ToString())
'End Using


#End Region

#End Region

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.Diagnostics
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms

Imports NativeMethods = DevCase.Win32.NativeMethods

Imports DevCase.Core.Application.Forms
Imports DevCase.Core.Application.Console
Imports DevCase.Win32
Imports DevCase.Win32.Delegates
Imports DevCase.Win32.Enums
Imports DevCase.Win32.Structures

#End Region

#Region " DevMessageBox "

' ReSharper disable once CheckNamespace

Namespace DevCase.UI.Dialogs

    ''' <summary>
    ''' Displays a message window, also known as a dialog box, which presents a message to the user.
    ''' <para></para>
    ''' It is a modal window, blocking other actions in the application until the user closes it.
    ''' </summary>
    '''
    ''' <remarks>
    ''' This class provides functionality similar to a <see cref="MessageBox"/>, with additional support 
    ''' for centering the dialog relative to its owner window, along with customization features.
    ''' </remarks>
    ''' 
    ''' <example> This is a code example that demonstrates static methods usage from a Windows Forms application.
    ''' <code language="VB">
    ''' Dim result As DialogResult = DevMessageBox.ShowStatic(Me, "Text", "Caption", "Details", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
    ''' Console.WriteLine(result.ToString())
    ''' </code>
    ''' </example>
    ''' 
    ''' <example> This is a code example that demonstrates static methods usage from a Console application.
    ''' <code language="VB">
    ''' Dim ownerWindow As New ConsoleWindowWrapper(UtilConsole.Handle)
    ''' DevCase.Win32.NativeMethods.SetForegroundWindow(ownerWindow.Handle)
    ''' 
    ''' Dim result As DialogResult = DevMessageBox.ShowStatic(Me, "Text", "Caption", "Details", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
    ''' Console.WriteLine(result.ToString())
    ''' </code>
    ''' </example>
    ''' 
    ''' <example> This is a code example that demonstrates instance methods usage from a Windows Forms application.
    ''' <code language="VB.NET">
    ''' Using msg As New DevMessageBox(owner:=Me, TimeSpan.FromSeconds(5))
    ''' 
    '''     Dim dlgresult As DialogResult = msg.Show("Question", "Title", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
    '''     Console.WriteLine(dlgresult.ToString())
    ''' End Using
    ''' </code>
    ''' </example>
    ''' 
    ''' <example> This is a code example that demonstrates instance methods usage from a Console application.
    ''' <code language="VB.NET">
    ''' Dim ownerWindow As New ConsoleWindowWrapper(UtilConsole.Handle)
    ''' DevCase.Win32.NativeMethods.SetForegroundWindow(ownerWindow.Handle)
    ''' 
    ''' Using msg As New DevMessageBox(ownerWindow, TimeSpan.FromSeconds(5))
    ''' 
    '''     Dim dlgresult As DialogResult = msg.Show("Question", "Title", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
    '''     Console.WriteLine(dlgresult.ToString())
    ''' End Using
    ''' </code>
    ''' </example>
    Public Class DevMessageBox : Implements IDisposable

#Region " Private Fields "

        ''' <summary>
        ''' Keeps track of the current amount of tries to find this <see cref="DevMessageBox"/> dialog.
        ''' </summary>
        Private enumWindowCurrentTries As Integer

        ''' <summary>
        ''' The maximum amount of tries to try find this <see cref="DevMessageBox"/> dialog.
        ''' </summary>
        Protected Const EnumWindowMaxTries As Integer = 100

        ''' <summary>
        ''' Number of seconds remaining before the message box closes automatically.
        ''' </summary>
        Private countdownSeconds As Integer

        ''' <summary>
        ''' Handle to the native countdown label window created via WinAPI.
        ''' <para></para>
        ''' Used to update the label text and manage its visibility during the timeout countdown.
        ''' </summary>
        Protected countdownLabelHandle As IntPtr

        ''' <summary>
        ''' A <see cref="System.Windows.Forms.Timer"/> used to decrement the countdown and update the label each second.
        ''' </summary>
        Protected WithEvents CountdownTimer As System.Windows.Forms.Timer

        ''' <summary>
        ''' A <see cref="System.Windows.Forms.Timer"/> that keeps track of <see cref="DevMessageBox.TimeOut"/> value to close this <see cref="DevMessageBox"/>.
        ''' </summary>
        Protected WithEvents DestroyTimer As System.Windows.Forms.Timer

        ''' <summary>
        ''' The messagebox main window handle (hwnd).
        ''' </summary>
        Protected DialogWindowHandle As IntPtr

        ''' <summary>
        ''' When the class is used from a console application, 
        ''' this value is used to keep track of the console size and position when the class is used from a console application.
        ''' </summary>
        Private consoleRect As NativeRectangle

        ''' <summary>
        ''' When the class is used from a console application, 
        ''' this value holds a transparent form that enables to enumerate thread windows to find the dialog window.
        ''' </summary>
        Private transparentForm As TransparentControlsForm

        ''' <summary>
        ''' Stores a handle to the previous window procedure (WndProc) for the dialog window, 
        ''' used to call the original window procedure after subclassing.
        ''' </summary>
        Private previousDlgWndProc As IntPtr

        ''' <summary>
        ''' Stores a handle to the native brush used to paint the dialog background.
        ''' </summary>
        Private dialogBackgroundBrush As IntPtr = IntPtr.Zero

        ''' <summary>
        ''' Stores a handle to the native brush used to paint the dialog footer background.
        ''' </summary>
        Private dialogFooterBackgroundBrush As IntPtr = IntPtr.Zero

        ''' <summary>
        ''' Stores a handle to the native brush used to paint the countdown label background.
        ''' </summary>
        Private countdownBrush As IntPtr = IntPtr.Zero

        ''' <summary>
        ''' The <see cref="Control"/> that owns this <see cref="DevMessageBox"/>.
        ''' <para></para>
        ''' This value is <c>null</c> if no <see cref="Control"/> has been passed to the <see cref="DevMessageBox"/> constructor.
        ''' </summary>
        Protected ownerControl As Control

#End Region

#Region " Properties "

        ''' <summary>
        ''' Gets the <see cref="IWin32Window"/> that owns this <see cref="DevMessageBox"/>.
        ''' <para></para>
        ''' This value is <c>null</c> if no <see cref="IWin32Window"/> has been passed to the <see cref="DevMessageBox"/> constructor.
        ''' </summary>
        '''
        ''' <value>
        ''' The the <see cref="IWin32Window"/> that owns this <see cref="DevMessageBox"/>; 
        ''' or <c>null</c> if no <see cref="IWin32Window"/> has been passed to the <see cref="DevMessageBox"/> constructor.
        ''' </value>
        Public ReadOnly Property OwnerWindow As IWin32Window
            Get
                If Me.ownerWindow_ IsNot Nothing Then
                    Return Me.ownerWindow_

                ElseIf Me.ownerControl IsNot Nothing Then
                    Return NativeWindow.FromHandle(Me.ownerControl.Handle)

                Else
                    Return Nothing

                End If
            End Get
        End Property
        ''' <summary>
        ''' (Backing Field)
        ''' <para></para>
        ''' The <see cref="IWin32Window"/> that owns this <see cref="DevMessageBox"/>.
        ''' <para></para>
        ''' This value is <c>null</c> if no <see cref="IWin32Window"/> has been passed to the <see cref="DevMessageBox"/> constructor.
        ''' </summary>
        Private ReadOnly ownerWindow_ As IWin32Window

        ''' <summary>
        ''' Gets or sets the time interval that this <see cref="DevMessageBox"/> will close automatically.
        ''' <para></para>
        ''' Minimum allowed value is 2000 total milliseconds (2 seconds).
        ''' <para></para>
        ''' Maximum allowed value is 359999000 total milliseconds (99h:59m:59s).
        ''' <para></para>
        ''' If this value is zero (<see cref="TimeSpan.Zero"/>) the <see cref="DevMessageBox"/> will not display a countdown.
        ''' <para></para>
        ''' Default value is <see cref="TimeSpan.Zero"/>.
        ''' </summary>
        Public Property TimeOut As TimeSpan
            Get
                Return Me.timeOut_
            End Get
            <DebuggerStepThrough>
            Set(value As TimeSpan)
                If value = Nothing Then
                    Me.timeOut_ = TimeSpan.Zero
                Else
                    Dim ms As Double = value.TotalMilliseconds
                    Const MinMs As Double = 2000      ' 2 seconds.
                    Const MaxMs As Double = 359999000 ' 99h:59m:59s.

                    If ms < MinMs Then
                        Throw New ArgumentOutOfRangeException(
                            NameOf(DevMessageBox.TimeOut), value,
                            $"Timeout value is out of range, it must be at least 2 seconds or higher."
                        )
                    End If

                    If ms > MaxMs Then
                        Throw New ArgumentOutOfRangeException(
                            NameOf(DevMessageBox.TimeOut), value,
                            $"Timeout value is out of range, total milliseconds can't exceed {MaxMs} (99h:59m:59s)."
                        )
                    End If
                End If

                timeOut_ = value
            End Set
        End Property

        ''' <summary>
        ''' (Backing Field)
        ''' <para></para>
        ''' The time interval that this <see cref="DevMessageBox"/> will close automatically.
        ''' </summary>
        Private timeOut_ As TimeSpan = TimeSpan.Zero

        ''' <summary>
        ''' Gets or sets the <see cref="WindowStylesEx"/> used to create the countdown label.
        ''' </summary>
        Public Property CountdownWindowStyles As WindowStylesEx = WindowStylesEx.Default

        ''' <summary>
        ''' Gets or sets the color used to paint the countdown label foreground.
        ''' </summary>
        Public Property CountdownForegroundColor As Color = SystemColors.ControlText

        ''' <summary>
        ''' Gets or sets the color used to paint the countdown label background.
        ''' </summary>
        Public Property CountdownBackgroundColor As Color = SystemColors.Window

        ''' <summary>
        ''' Gets or sets the background color used to paint the dialog window and its child controls.
        ''' <para></para>
        ''' Default value is <see cref="Color.Empty"/>, which means no custom background is applied.
        ''' </summary>
        Public Property DialogBackgroundColor As Color = Color.Empty

        ''' <summary>
        ''' Gets or sets the background color used to paint the dialog window and its child controls.
        ''' <para></para>
        ''' Default value is <see cref="Color.Empty"/>, which means no custom background is applied.
        ''' </summary>
        Public Property DialogFooterBackgroundColor As Color = Color.Empty

        ''' <summary>
        ''' Gets or sets the foreground (text) color used to paint the dialog window's child controls.
        ''' <para></para>
        ''' Default value is <see cref="Color.Empty"/>, which means no custom foreground is applied.
        ''' </summary>
        Public Property DialogForegroundColor As Color = Color.Empty

        ''' <summary>
        ''' Gets or sets a value indicating whether a dark title bar should be used
        ''' when supported by the operating system.
        ''' </summary>
        ''' 
        ''' <value>
        ''' <c>True</c> to enable the dark title bar style; otherwise, <c>False</c>.
        ''' </value>
        ''' <remarks>
        ''' 
        ''' This property applies only on supported Windows versions that provide
        ''' native dark title bar support. If the operating system does not support
        ''' this feature, the setting is ignored and the default system title bar
        ''' appearance is used.
        ''' </remarks>
        Public Property UseDarkTitleBar As Boolean = False

#End Region

#Region " Constructors "

        ''' <summary>
        ''' Initializes a new instance of the <see cref="DevMessageBox"/> class.
        ''' </summary>
        '''
        ''' <param name="owner">
        ''' The <see cref="IWin32Window"/> that acts as the owner of this <see cref="DevMessageBox"/>.
        ''' <para></para>
        ''' Assigning an owner enables modal behavior, blocking input until this message box closes,
        ''' and ensures the message box is centered relative to the owner's bounds.
        ''' <para></para>
        ''' This value can be <c>null</c>, in such case the message box will behave as a standard modal dialog window 
        ''' centered to the current screen bounds.
        ''' </param>
        ''' 
        ''' <param name="timeOut">
        ''' The time interval that this <see cref="DevMessageBox"/> will close automatically.
        ''' <para></para>
        ''' Minimum allowed value is 2000 total milliseconds (2 seconds).
        ''' <para></para>
        ''' Maximum allowed value is 359999000 total milliseconds (99h:59m:59s).
        ''' <para></para>
        ''' If this value is zero (<see cref="TimeSpan.Zero"/>) the <see cref="DevMessageBox"/> will not display a countdown.
        ''' <para></para>
        ''' Default value is <see cref="TimeSpan.Zero"/>.
        ''' </param>
        <DebuggerStepThrough>
        Public Sub New(owner As IWin32Window, Optional timeOut As TimeSpan = Nothing)

            Me.TimeOut = timeOut

            If TypeOf owner Is Form Then
                Me.ownerControl = DirectCast(owner, Form)
            ElseIf TypeOf owner Is Control Then
                Me.ownerControl = DirectCast(owner, Control)
            End If
            Me.ownerWindow_ = owner
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the <see cref="DevMessageBox"/> class.
        ''' </summary>
        '''
        ''' <param name="owner">
        ''' The <see cref="Form"/> that acts as the owner of this <see cref="DevMessageBox"/>.
        ''' <para></para>
        ''' Assigning an owner enables modal behavior, blocking input until this message box closes,
        ''' and ensures the message box is centered relative to the owner's bounds.
        ''' <para></para>
        ''' This value can be <c>null</c>, in such case the message box will behave as a standard modal dialog window 
        ''' centered to the current screen bounds.
        ''' </param>
        ''' 
        ''' <param name="timeOut">
        ''' The time interval that this <see cref="DevMessageBox"/> will close automatically.
        ''' <para></para>
        ''' Minimum allowed value is 2000 total milliseconds (2 seconds).
        ''' <para></para>
        ''' Maximum allowed value is 359999000 total milliseconds (99h:59m:59s).
        ''' <para></para>
        ''' If this value is zero (<see cref="TimeSpan.Zero"/>) the <see cref="DevMessageBox"/> will not display a countdown.
        ''' <para></para>
        ''' Default value is <see cref="TimeSpan.Zero"/>.
        ''' </param>
        <DebuggerStepThrough>
        Public Sub New(owner As Form, Optional timeOut As TimeSpan = Nothing)

            Me.New(If(owner Is Nothing, Nothing, NativeWindow.FromHandle(owner.Handle)), timeOut)
            Me.ownerControl = owner
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the <see cref="DevMessageBox"/> class.
        ''' </summary>
        '''
        ''' <param name="owner">
        ''' The <see cref="Control"/> that acts as the owner of this <see cref="DevMessageBox"/>.
        ''' <para></para>
        ''' Assigning an owner enables modal behavior, blocking input until this message box closes,
        ''' and ensures the message box is centered relative to the owner's bounds.
        ''' <para></para>
        ''' This value can be <c>null</c>, in such case the message box will behave as a standard modal dialog window 
        ''' centered to the current screen bounds.
        ''' </param>
        ''' 
        ''' <param name="timeOut">
        ''' The time interval that this <see cref="DevMessageBox"/> will close automatically.
        ''' <para></para>
        ''' Minimum allowed value is 2000 total milliseconds (2 seconds).
        ''' <para></para>
        ''' Maximum allowed value is 359999000 total milliseconds (99h:59m:59s).
        ''' <para></para>
        ''' If this value is zero (<see cref="TimeSpan.Zero"/>) the <see cref="DevMessageBox"/> will not display a countdown.
        ''' <para></para>
        ''' Default value is <see cref="TimeSpan.Zero"/>.
        ''' </param>
        <DebuggerStepThrough>
        Public Sub New(owner As Control, Optional timeOut As TimeSpan = Nothing)

            Me.New(If(owner Is Nothing, Nothing, NativeWindow.FromHandle(owner.Handle)), timeOut)
            Me.ownerControl = owner
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the <see cref="DevMessageBox"/> class.
        ''' </summary>
        '''
        ''' <param name="owner">
        ''' The <see cref="ConsoleWindowWrapper"/> that acts as the owner of this <see cref="DevMessageBox"/>.
        ''' <para></para>
        ''' Assigning an owner enables modal behavior, blocking input until this message box closes,
        ''' and ensures the message box is centered relative to the owner's bounds.
        ''' <para></para>
        ''' This value can be <c>null</c>, in such case the message box will behave as a standard modal dialog window 
        ''' centered to the current screen bounds.
        ''' </param>
        ''' 
        ''' <param name="timeOut">
        ''' The time interval that this <see cref="DevMessageBox"/> will close automatically.
        ''' <para></para>
        ''' Minimum allowed value is 2000 total milliseconds (2 seconds).
        ''' <para></para>
        ''' Maximum allowed value is 359999000 total milliseconds (99h:59m:59s).
        ''' <para></para>
        ''' If this value is zero (<see cref="TimeSpan.Zero"/>) the <see cref="DevMessageBox"/> will not display a countdown.
        ''' <para></para>
        ''' Default value is <see cref="TimeSpan.Zero"/>.
        ''' </param>
        <DebuggerStepThrough>
        Public Sub New(owner As ConsoleWindowWrapper, Optional timeOut As TimeSpan = Nothing)

            Me.New(DirectCast(owner, IWin32Window), timeOut)
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the <see cref="DevMessageBox"/> class 
        ''' without assigning an owner window.
        ''' </summary>
        <DebuggerStepThrough>
        Public Sub New()

            Me.New(DirectCast(Nothing, IWin32Window), TimeSpan.Zero)
        End Sub

#End Region

#Region " Public Methods "

        ''' <summary>
        ''' Displays a message box with specified text.
        ''' </summary>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' </returns>
        <DebuggerStepThrough>
        Public Overridable Function Show(text As String) As DialogResult

            Me.CleanUp()
            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(Me.OwnerWindow, text))
        End Function

        ''' <summary>
        ''' Displays a message box with specified text and caption.
        ''' </summary>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="caption">
        ''' The text to display in the title bar of the message box.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' </returns>
        <DebuggerStepThrough>
        Public Overridable Function Show(text As String, caption As String) As DialogResult

            Me.CleanUp()
            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(Me.OwnerWindow, text, caption))
        End Function

        ''' <summary>
        ''' Displays a message box with specified text, caption, and buttons.
        ''' </summary>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="caption">
        ''' The text to display in the title bar of the message box.
        ''' </param>
        ''' 
        ''' <param name="buttons">
        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' <para></para>
        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
        ''' </returns>
        <DebuggerStepThrough>
        Public Overridable Function Show(text As String, caption As String, buttons As MessageBoxButtons) As DialogResult

            Me.CleanUp()
            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(Me.OwnerWindow, text, caption, buttons))
        End Function

        ''' <summary>
        ''' Displays a message box with specified text, caption, buttons, and icon.
        ''' </summary>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="caption">
        ''' The text to display in the title bar of the message box.
        ''' </param>
        ''' 
        ''' <param name="buttons">
        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="icon">
        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' <para></para>
        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
        ''' </returns>
        <DebuggerStepThrough>
        Public Overridable Function Show(text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon) As DialogResult

            Me.CleanUp()
            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(Me.OwnerWindow, text, caption, buttons, icon))
        End Function

        ''' <summary>
        ''' Displays a message box with the specified text, caption, buttons, icon, and default button.
        ''' </summary>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="caption">
        ''' The text to display in the title bar of the message box.
        ''' </param>
        ''' 
        ''' <param name="buttons">
        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="icon">
        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="defaultButton">
        ''' One of the <see cref="MessageBoxDefaultButton"/> values that specifies the default button for the message box.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' <para></para>
        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
        ''' </returns>
        <DebuggerStepThrough>
        Public Overridable Function Show(text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon, defaultButton As MessageBoxDefaultButton) As DialogResult

            Me.CleanUp()
            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(Me.OwnerWindow, text, caption, buttons, icon, defaultButton))
        End Function

        ''' <summary>
        ''' Displays a message box with the specified text, caption, buttons, icon, default button, and options.
        ''' </summary>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="caption">
        ''' The text to display in the title bar of the message box.
        ''' </param>
        ''' 
        ''' <param name="buttons">
        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="icon">
        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="defaultButton">
        ''' One of the <see cref="MessageBoxDefaultButton"/> values that specifies the default button for the message box.
        ''' </param>
        ''' 
        ''' <param name="options">
        ''' One of the <see cref="MessageBoxOptions"/> values that specifies which display and association options will be used for 
        ''' the message box. 
        ''' <para></para>
        ''' You may pass in 0 if you wish to use the defaults.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' <para></para>
        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
        ''' </returns>
        <DebuggerStepThrough>
        Public Overridable Function Show(text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon, defaultButton As MessageBoxDefaultButton, options As MessageBoxOptions) As DialogResult

            Me.CleanUp()
            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(Me.OwnerWindow, text, caption, buttons, icon, defaultButton, options))
        End Function

        ''' <summary>
        ''' Displays a message box with the specified text, caption, buttons, icon, default button, options, and Help button, 
        ''' using the specified Help file.
        ''' </summary>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="caption">
        ''' The text to display in the title bar of the message box.
        ''' </param>
        ''' 
        ''' <param name="buttons">
        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="icon">
        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="defaultButton">
        ''' One of the <see cref="MessageBoxDefaultButton"/> values that specifies the default button for the message box.
        ''' </param>
        ''' 
        ''' <param name="options">
        ''' One of the <see cref="MessageBoxOptions"/> values that specifies which display and association options will be used for 
        ''' the message box. 
        ''' <para></para>
        ''' You may pass in 0 if you wish to use the defaults.
        ''' </param>
        ''' 
        ''' <param name="helpFilePath">
        ''' The path and name of the Help file to display when the user clicks the Help button.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' <para></para>
        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
        ''' </returns>
        <DebuggerStepThrough>
        Public Overridable Function Show(text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon, defaultButton As MessageBoxDefaultButton, options As MessageBoxOptions, helpFilePath As String) As DialogResult

            Me.CleanUp()
            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(Me.OwnerWindow, text, caption, buttons, icon, defaultButton, options, helpFilePath))
        End Function

        ''' <summary>
        ''' Displays a message box with the specified text, caption, buttons, icon, default button, options, and Help button.
        ''' </summary>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="caption">
        ''' The text to display in the title bar of the message box.
        ''' </param>
        ''' 
        ''' <param name="buttons">
        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="icon">
        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="defaultButton">
        ''' One of the <see cref="MessageBoxDefaultButton"/> values that specifies the default button for the message box.
        ''' </param>
        ''' 
        ''' <param name="options">
        ''' One of the <see cref="MessageBoxOptions"/> values that specifies which display and association options will be used for 
        ''' the message box. 
        ''' <para></para>
        ''' You may pass in 0 if you wish to use the defaults.
        ''' </param>
        ''' 
        ''' <param name="displayHelpButton">
        ''' <see langword="True"/> to show the Help button; otherwise, false. The default is <see langword="False"/>.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' <para></para>
        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
        ''' </returns>
        <DebuggerStepThrough>
        Public Overridable Function Show(text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon, defaultButton As MessageBoxDefaultButton, options As MessageBoxOptions, displayHelpButton As Boolean) As DialogResult

            Me.CleanUp()
            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(text, caption, buttons, icon, defaultButton, options, displayHelpButton))
        End Function

        ''' <summary>
        ''' Displays a message box with the specified text, caption, buttons, icon, default button, options, and Help button, 
        ''' using the specified Help file and HelpNavigator.
        ''' </summary>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="caption">
        ''' The text to display in the title bar of the message box.
        ''' </param>
        ''' 
        ''' <param name="buttons">
        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="icon">
        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="defaultButton">
        ''' One of the <see cref="MessageBoxDefaultButton"/> values that specifies the default button for the message box.
        ''' </param>
        ''' 
        ''' <param name="options">
        ''' One of the <see cref="MessageBoxOptions"/> values that specifies which display and association options will be used for 
        ''' the message box. 
        ''' <para></para>
        ''' You may pass in 0 if you wish to use the defaults.
        ''' </param>
        ''' 
        ''' <param name="helpFilePath">
        ''' The path and name of the Help file to display when the user clicks the Help button.
        ''' </param>
        ''' 
        ''' <param name="navigator">
        ''' One of the <see cref="HelpNavigator"/> values.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' <para></para>
        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
        ''' </returns>
        <DebuggerStepThrough>
        Public Overridable Function Show(text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon, defaultButton As MessageBoxDefaultButton, options As MessageBoxOptions, helpFilePath As String, navigator As HelpNavigator) As DialogResult

            Me.CleanUp()
            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(Me.OwnerWindow, text, caption, buttons, icon, defaultButton, options, helpFilePath, navigator))
        End Function

        ''' <summary>
        ''' Displays a message box with the specified text, caption, buttons, icon, default button, options, and Help button, 
        ''' using the specified Help file and Help keyword.
        ''' </summary>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="caption">
        ''' The text to display in the title bar of the message box.
        ''' </param>
        ''' 
        ''' <param name="buttons">
        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="icon">
        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="defaultButton">
        ''' One of the <see cref="MessageBoxDefaultButton"/> values that specifies the default button for the message box.
        ''' </param>
        ''' 
        ''' <param name="options">
        ''' One of the <see cref="MessageBoxOptions"/> values that specifies which display and association options will be used for 
        ''' the message box. 
        ''' <para></para>
        ''' You may pass in 0 if you wish to use the defaults.
        ''' </param>
        ''' 
        ''' <param name="helpFilePath">
        ''' The path and name of the Help file to display when the user clicks the Help button.
        ''' </param>
        ''' 
        ''' <param name="keyword">
        ''' The Help keyword to display when the user clicks the Help button.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' <para></para>
        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
        ''' </returns>
        <DebuggerStepThrough>
        Public Overridable Function Show(text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon, defaultButton As MessageBoxDefaultButton, options As MessageBoxOptions, helpFilePath As String, keyword As String) As DialogResult

            Me.CleanUp()
            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(Me.OwnerWindow, text, caption, buttons, icon, defaultButton, options, helpFilePath, keyword))
        End Function

        ''' <summary>
        ''' Displays a message box with the specified text, caption, buttons, icon, default button, options, and Help button, 
        ''' using the specified Help file, HelpNavigator, and Help topic.
        ''' </summary>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="caption">
        ''' The text to display in the title bar of the message box.
        ''' </param>
        ''' 
        ''' <param name="buttons">
        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="icon">
        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="defaultButton">
        ''' One of the <see cref="MessageBoxDefaultButton"/> values that specifies the default button for the message box.
        ''' </param>
        ''' 
        ''' <param name="options">
        ''' One of the <see cref="MessageBoxOptions"/> values that specifies which display and association options will be used for 
        ''' the message box. 
        ''' <para></para>
        ''' You may pass in 0 if you wish to use the defaults.
        ''' </param>
        ''' 
        ''' <param name="helpFilePath">
        ''' The path and name of the Help file to display when the user clicks the Help button.
        ''' </param>
        ''' 
        ''' <param name="navigator">
        ''' One of the <see cref="HelpNavigator"/> values.
        ''' </param>
        ''' 
        ''' <param name="param">
        ''' The numeric ID of the Help topic to display when the user clicks the Help button.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' <para></para>
        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
        ''' </returns>
        <DebuggerStepThrough>
        Public Overridable Function Show(text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon, defaultButton As MessageBoxDefaultButton, options As MessageBoxOptions, helpFilePath As String, navigator As HelpNavigator, param As Object) As DialogResult

            Me.CleanUp()
            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(Me.OwnerWindow, text, caption, buttons, icon, defaultButton, options, helpFilePath, navigator, param))
        End Function

#End Region

#Region " Public Static Methods "

        ''' <summary>
        ''' Displays a message box with specified text.
        ''' </summary>
        '''
        ''' <remarks>
        ''' This is a convenience static method that creates a temporary <see cref="DevMessageBox"/> instance,
        ''' displays the dialog, and disposes the instance automatically when the dialog is closed.
        ''' </remarks>
        ''' 
        ''' <param name="owner">
        ''' The <see cref="IWin32Window"/> that acts as the owner of this <see cref="DevMessageBox"/>.
        ''' <para></para>
        ''' Assigning an owner enables modal behavior, blocking input until this message box closes,
        ''' and ensures the message box is centered relative to the owner's bounds.
        ''' <para></para>
        ''' This value can be <c>null</c>, in such case the message box will behave as a standard modal dialog window 
        ''' centered to the current screen bounds.
        ''' </param>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' </returns>
        <DebuggerStepThrough>
        Public Shared Function ShowStatic(owner As IWin32Window, text As String) As DialogResult

            Using msg As New DevMessageBox(owner)
                Return msg.Show(text:=text)
            End Using
        End Function

        ''' <summary>
        ''' Displays a message box with specified text and caption.
        ''' </summary>
        '''
        ''' <remarks>
        ''' This is a convenience static method that creates a temporary <see cref="DevMessageBox"/> instance,
        ''' displays the dialog, and disposes the instance automatically when the dialog is closed.
        ''' </remarks>
        ''' 
        ''' <param name="owner">
        ''' The <see cref="IWin32Window"/> that acts as the owner of this <see cref="DevMessageBox"/>.
        ''' <para></para>
        ''' Assigning an owner enables modal behavior, blocking input until this message box closes,
        ''' and ensures the message box is centered relative to the owner's bounds.
        ''' <para></para>
        ''' This value can be <c>null</c>, in such case the message box will behave as a standard modal dialog window 
        ''' centered to the current screen bounds.
        ''' </param>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="caption">
        ''' The text to display in the title bar of the message box.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' </returns>
        <DebuggerStepThrough>
        Public Shared Function ShowStatic(owner As IWin32Window, text As String, caption As String) As DialogResult

            Using msg As New DevMessageBox(owner)
                Return msg.Show(text:=text, caption:=caption)
            End Using
        End Function

        ''' <summary>
        ''' Displays a message box with specified text, caption, and buttons.
        ''' </summary>
        '''
        ''' <remarks>
        ''' This is a convenience static method that creates a temporary <see cref="DevMessageBox"/> instance,
        ''' displays the dialog, and disposes the instance automatically when the dialog is closed.
        ''' </remarks>
        ''' 
        ''' <param name="owner">
        ''' The <see cref="IWin32Window"/> that acts as the owner of this <see cref="DevMessageBox"/>.
        ''' <para></para>
        ''' Assigning an owner enables modal behavior, blocking input until this message box closes,
        ''' and ensures the message box is centered relative to the owner's bounds.
        ''' <para></para>
        ''' This value can be <c>null</c>, in such case the message box will behave as a standard modal dialog window 
        ''' centered to the current screen bounds.
        ''' </param>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="caption">
        ''' The text to display in the title bar of the message box.
        ''' </param>
        ''' 
        ''' <param name="buttons">
        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' <para></para>
        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
        ''' </returns>
        <DebuggerStepThrough>
        Public Shared Function ShowStatic(owner As IWin32Window, text As String, caption As String, buttons As MessageBoxButtons) As DialogResult

            Using msg As New DevMessageBox(owner)
                Return msg.Show(text:=text, caption:=caption, buttons:=buttons)
            End Using
        End Function

        ''' <summary>
        ''' Displays a message box with specified text, caption, buttons, and icon.
        ''' </summary>
        '''
        ''' <remarks>
        ''' This is a convenience static method that creates a temporary <see cref="DevMessageBox"/> instance,
        ''' displays the dialog, and disposes the instance automatically when the dialog is closed.
        ''' </remarks>
        ''' 
        ''' <param name="owner">
        ''' The <see cref="IWin32Window"/> that acts as the owner of this <see cref="DevMessageBox"/>.
        ''' <para></para>
        ''' Assigning an owner enables modal behavior, blocking input until this message box closes,
        ''' and ensures the message box is centered relative to the owner's bounds.
        ''' <para></para>
        ''' This value can be <c>null</c>, in such case the message box will behave as a standard modal dialog window 
        ''' centered to the current screen bounds.
        ''' </param>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="caption">
        ''' The text to display in the title bar of the message box.
        ''' </param>
        ''' 
        ''' <param name="buttons">
        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="icon">
        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' <para></para>
        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
        ''' </returns>
        <DebuggerStepThrough>
        Public Shared Function ShowStatic(owner As IWin32Window, text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon) As DialogResult

            Using msg As New DevMessageBox(owner)
                Return msg.Show(text:=text, caption:=caption, buttons:=buttons, icon:=icon)
            End Using
        End Function

        ''' <summary>
        ''' Displays a message box with the specified text, caption, buttons, icon, and default button.
        ''' </summary>
        '''
        ''' <remarks>
        ''' This is a convenience static method that creates a temporary <see cref="DevMessageBox"/> instance,
        ''' displays the dialog, and disposes the instance automatically when the dialog is closed.
        ''' </remarks>
        ''' 
        ''' <param name="owner">
        ''' The <see cref="IWin32Window"/> that acts as the owner of this <see cref="DevMessageBox"/>.
        ''' <para></para>
        ''' Assigning an owner enables modal behavior, blocking input until this message box closes,
        ''' and ensures the message box is centered relative to the owner's bounds.
        ''' <para></para>
        ''' This value can be <c>null</c>, in such case the message box will behave as a standard modal dialog window 
        ''' centered to the current screen bounds.
        ''' </param>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="caption">
        ''' The text to display in the title bar of the message box.
        ''' </param>
        ''' 
        ''' <param name="buttons">
        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="icon">
        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="defaultButton">
        ''' One of the <see cref="MessageBoxDefaultButton"/> values that specifies the default button for the message box.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' <para></para>
        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
        ''' </returns>
        <DebuggerStepThrough>
        Public Shared Function ShowStatic(owner As IWin32Window, text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon, defaultButton As MessageBoxDefaultButton) As DialogResult

            Using msg As New DevMessageBox(owner)
                Return msg.Show(text:=text, caption:=caption, buttons:=buttons, icon:=icon, defaultButton:=defaultButton)
            End Using
        End Function

        ''' <summary>
        ''' Displays a message box with the specified text, caption, buttons, icon, default button, and options.
        ''' </summary>
        '''
        ''' <remarks>
        ''' This is a convenience static method that creates a temporary <see cref="DevMessageBox"/> instance,
        ''' displays the dialog, and disposes the instance automatically when the dialog is closed.
        ''' </remarks>
        ''' 
        ''' <param name="owner">
        ''' The <see cref="IWin32Window"/> that acts as the owner of this <see cref="DevMessageBox"/>.
        ''' <para></para>
        ''' Assigning an owner enables modal behavior, blocking input until this message box closes,
        ''' and ensures the message box is centered relative to the owner's bounds.
        ''' <para></para>
        ''' This value can be <c>null</c>, in such case the message box will behave as a standard modal dialog window 
        ''' centered to the current screen bounds.
        ''' </param>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="caption">
        ''' The text to display in the title bar of the message box.
        ''' </param>
        ''' 
        ''' <param name="buttons">
        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="icon">
        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="defaultButton">
        ''' One of the <see cref="MessageBoxDefaultButton"/> values that specifies the default button for the message box.
        ''' </param>
        ''' 
        ''' <param name="options">
        ''' One of the <see cref="MessageBoxOptions"/> values that specifies which display and association options will be used for 
        ''' the message box. 
        ''' <para></para>
        ''' You may pass in 0 if you wish to use the defaults.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' <para></para>
        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
        ''' </returns>
        <DebuggerStepThrough>
        Public Shared Function ShowStatic(owner As IWin32Window, text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon, defaultButton As MessageBoxDefaultButton, options As MessageBoxOptions) As DialogResult

            Using msg As New DevMessageBox(owner)
                Return msg.Show(text:=text, caption:=caption, buttons:=buttons, icon:=icon, defaultButton:=defaultButton, options:=options)
            End Using
        End Function

        ''' <summary>
        ''' Displays a message box with the specified text, caption, buttons, icon, default button, options, and Help button, 
        ''' using the specified Help file.
        ''' </summary>
        '''
        ''' <remarks>
        ''' This is a convenience static method that creates a temporary <see cref="DevMessageBox"/> instance,
        ''' displays the dialog, and disposes the instance automatically when the dialog is closed.
        ''' </remarks>
        ''' 
        ''' <param name="owner">
        ''' The <see cref="IWin32Window"/> that acts as the owner of this <see cref="DevMessageBox"/>.
        ''' <para></para>
        ''' Assigning an owner enables modal behavior, blocking input until this message box closes,
        ''' and ensures the message box is centered relative to the owner's bounds.
        ''' <para></para>
        ''' This value can be <c>null</c>, in such case the message box will behave as a standard modal dialog window 
        ''' centered to the current screen bounds.
        ''' </param>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="caption">
        ''' The text to display in the title bar of the message box.
        ''' </param>
        ''' 
        ''' <param name="buttons">
        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="icon">
        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="defaultButton">
        ''' One of the <see cref="MessageBoxDefaultButton"/> values that specifies the default button for the message box.
        ''' </param>
        ''' 
        ''' <param name="options">
        ''' One of the <see cref="MessageBoxOptions"/> values that specifies which display and association options will be used for 
        ''' the message box. 
        ''' <para></para>
        ''' You may pass in 0 if you wish to use the defaults.
        ''' </param>
        ''' 
        ''' <param name="helpFilePath">
        ''' The path and name of the Help file to display when the user clicks the Help button.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' <para></para>
        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
        ''' </returns>
        <DebuggerStepThrough>
        Public Shared Function ShowStatic(owner As IWin32Window, text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon, defaultButton As MessageBoxDefaultButton, options As MessageBoxOptions, helpFilePath As String) As DialogResult

            Using msg As New DevMessageBox(owner)
                Return msg.Show(text:=text, caption:=caption, buttons:=buttons, icon:=icon, defaultButton:=defaultButton, options:=options, helpFilePath:=helpFilePath)
            End Using
        End Function

        ''' <summary>
        ''' Displays a message box with the specified text, caption, buttons, icon, default button, options, and Help button.
        ''' </summary>
        '''
        ''' <remarks>
        ''' This is a convenience static method that creates a temporary <see cref="DevMessageBox"/> instance,
        ''' displays the dialog, and disposes the instance automatically when the dialog is closed.
        ''' </remarks>
        ''' 
        ''' <param name="owner">
        ''' The <see cref="IWin32Window"/> that acts as the owner of this <see cref="DevMessageBox"/>.
        ''' <para></para>
        ''' Assigning an owner enables modal behavior, blocking input until this message box closes,
        ''' and ensures the message box is centered relative to the owner's bounds.
        ''' <para></para>
        ''' This value can be <c>null</c>, in such case the message box will behave as a standard modal dialog window 
        ''' centered to the current screen bounds.
        ''' </param>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="caption">
        ''' The text to display in the title bar of the message box.
        ''' </param>
        ''' 
        ''' <param name="buttons">
        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="icon">
        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="defaultButton">
        ''' One of the <see cref="MessageBoxDefaultButton"/> values that specifies the default button for the message box.
        ''' </param>
        ''' 
        ''' <param name="options">
        ''' One of the <see cref="MessageBoxOptions"/> values that specifies which display and association options will be used for 
        ''' the message box. 
        ''' <para></para>
        ''' You may pass in 0 if you wish to use the defaults.
        ''' </param>
        ''' 
        ''' <param name="displayHelpButton">
        ''' <see langword="True"/> to show the Help button; otherwise, false. The default is <see langword="False"/>.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' <para></para>
        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
        ''' </returns>
        <DebuggerStepThrough>
        Public Shared Function ShowStatic(owner As IWin32Window, text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon, defaultButton As MessageBoxDefaultButton, options As MessageBoxOptions, displayHelpButton As Boolean) As DialogResult

            Using msg As New DevMessageBox(owner)
                Return msg.Show(text:=text, caption:=caption, buttons:=buttons, icon:=icon, defaultButton:=defaultButton, options:=options, displayHelpButton:=displayHelpButton)
            End Using
        End Function

        ''' <summary>
        ''' Displays a message box with the specified text, caption, buttons, icon, default button, options, and Help button, 
        ''' using the specified Help file and HelpNavigator.
        ''' </summary>
        '''
        ''' <remarks>
        ''' This is a convenience static method that creates a temporary <see cref="DevMessageBox"/> instance,
        ''' displays the dialog, and disposes the instance automatically when the dialog is closed.
        ''' </remarks>
        ''' 
        ''' <param name="owner">
        ''' The <see cref="IWin32Window"/> that acts as the owner of this <see cref="DevMessageBox"/>.
        ''' <para></para>
        ''' Assigning an owner enables modal behavior, blocking input until this message box closes,
        ''' and ensures the message box is centered relative to the owner's bounds.
        ''' <para></para>
        ''' This value can be <c>null</c>, in such case the message box will behave as a standard modal dialog window 
        ''' centered to the current screen bounds.
        ''' </param>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="caption">
        ''' The text to display in the title bar of the message box.
        ''' </param>
        ''' 
        ''' <param name="buttons">
        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="icon">
        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="defaultButton">
        ''' One of the <see cref="MessageBoxDefaultButton"/> values that specifies the default button for the message box.
        ''' </param>
        ''' 
        ''' <param name="options">
        ''' One of the <see cref="MessageBoxOptions"/> values that specifies which display and association options will be used for 
        ''' the message box. 
        ''' <para></para>
        ''' You may pass in 0 if you wish to use the defaults.
        ''' </param>
        ''' 
        ''' <param name="helpFilePath">
        ''' The path and name of the Help file to display when the user clicks the Help button.
        ''' </param>
        ''' 
        ''' <param name="navigator">
        ''' One of the <see cref="HelpNavigator"/> values.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' <para></para>
        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
        ''' </returns>
        <DebuggerStepThrough>
        Public Shared Function ShowStatic(owner As IWin32Window, text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon, defaultButton As MessageBoxDefaultButton, options As MessageBoxOptions, helpFilePath As String, navigator As HelpNavigator) As DialogResult

            Using msg As New DevMessageBox(owner)
                Return msg.Show(text:=text, caption:=caption, buttons:=buttons, icon:=icon, defaultButton:=defaultButton, options:=options, helpFilePath:=helpFilePath, navigator:=navigator)
            End Using
        End Function

        ''' <summary>
        ''' Displays a message box with the specified text, caption, buttons, icon, default button, options, and Help button, 
        ''' using the specified Help file and Help keyword.
        ''' </summary>
        '''
        ''' <remarks>
        ''' This is a convenience static method that creates a temporary <see cref="DevMessageBox"/> instance,
        ''' displays the dialog, and disposes the instance automatically when the dialog is closed.
        ''' </remarks>
        ''' 
        ''' <param name="owner">
        ''' The <see cref="IWin32Window"/> that acts as the owner of this <see cref="DevMessageBox"/>.
        ''' <para></para>
        ''' Assigning an owner enables modal behavior, blocking input until this message box closes,
        ''' and ensures the message box is centered relative to the owner's bounds.
        ''' <para></para>
        ''' This value can be <c>null</c>, in such case the message box will behave as a standard modal dialog window 
        ''' centered to the current screen bounds.
        ''' </param>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="caption">
        ''' The text to display in the title bar of the message box.
        ''' </param>
        ''' 
        ''' <param name="buttons">
        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="icon">
        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="defaultButton">
        ''' One of the <see cref="MessageBoxDefaultButton"/> values that specifies the default button for the message box.
        ''' </param>
        ''' 
        ''' <param name="options">
        ''' One of the <see cref="MessageBoxOptions"/> values that specifies which display and association options will be used for 
        ''' the message box. 
        ''' <para></para>
        ''' You may pass in 0 if you wish to use the defaults.
        ''' </param>
        ''' 
        ''' <param name="helpFilePath">
        ''' The path and name of the Help file to display when the user clicks the Help button.
        ''' </param>
        ''' 
        ''' <param name="keyword">
        ''' The Help keyword to display when the user clicks the Help button.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' <para></para>
        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
        ''' </returns>
        <DebuggerStepThrough>
        Public Shared Function ShowStatic(owner As IWin32Window, text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon, defaultButton As MessageBoxDefaultButton, options As MessageBoxOptions, helpFilePath As String, keyword As String) As DialogResult

            Using msg As New DevMessageBox(owner)
                Return msg.Show(text:=text, caption:=caption, buttons:=buttons, icon:=icon, defaultButton:=defaultButton, options:=options, helpFilePath:=helpFilePath, keyword:=keyword)
            End Using
        End Function

        ''' <summary>
        ''' Displays a message box with the specified text, caption, buttons, icon, default button, options, and Help button, 
        ''' using the specified Help file, HelpNavigator, and Help topic.
        ''' </summary>
        '''
        ''' <remarks>
        ''' This is a convenience static method that creates a temporary <see cref="DevMessageBox"/> instance,
        ''' displays the dialog, and disposes the instance automatically when the dialog is closed.
        ''' </remarks>
        ''' 
        ''' <param name="owner">
        ''' The <see cref="IWin32Window"/> that acts as the owner of this <see cref="DevMessageBox"/>.
        ''' <para></para>
        ''' Assigning an owner enables modal behavior, blocking input until this message box closes,
        ''' and ensures the message box is centered relative to the owner's bounds.
        ''' <para></para>
        ''' This value can be <c>null</c>, in such case the message box will behave as a standard modal dialog window 
        ''' centered to the current screen bounds.
        ''' </param>
        '''
        ''' <param name="text">
        ''' The text to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="caption">
        ''' The text to display in the title bar of the message box.
        ''' </param>
        ''' 
        ''' <param name="buttons">
        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="icon">
        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
        ''' </param>
        ''' 
        ''' <param name="defaultButton">
        ''' One of the <see cref="MessageBoxDefaultButton"/> values that specifies the default button for the message box.
        ''' </param>
        ''' 
        ''' <param name="options">
        ''' One of the <see cref="MessageBoxOptions"/> values that specifies which display and association options will be used for 
        ''' the message box. 
        ''' <para></para>
        ''' You may pass in 0 if you wish to use the defaults.
        ''' </param>
        ''' 
        ''' <param name="helpFilePath">
        ''' The path and name of the Help file to display when the user clicks the Help button.
        ''' </param>
        ''' 
        ''' <param name="navigator">
        ''' One of the <see cref="HelpNavigator"/> values.
        ''' </param>
        ''' 
        ''' <param name="param">
        ''' The numeric ID of the Help topic to display when the user clicks the Help button.
        ''' </param>
        '''
        ''' <returns>
        ''' One of the <see cref="DialogResult"/> values.
        ''' <para></para>
        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
        ''' </returns>
        <DebuggerStepThrough>
        Public Shared Function ShowStatic(owner As IWin32Window, text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon, defaultButton As MessageBoxDefaultButton, options As MessageBoxOptions, helpFilePath As String, navigator As HelpNavigator, param As Object) As DialogResult

            Using msg As New DevMessageBox(owner)
                Return msg.Show(text:=text, caption:=caption, buttons:=buttons, icon:=icon, defaultButton:=defaultButton, options:=options, helpFilePath:=helpFilePath, navigator:=navigator, param:=param)
            End Using
        End Function

#End Region

#Region " Private Methods "

        ''' <summary>
        ''' Attempts to locate the message box dialog window by enumerating the thread's windows,
        ''' and initiates a retry mechanism if the dialog is not found immediately.
        ''' <para></para>
        ''' Also starts countdown and destruction timers if <see cref="DevMessageBox.TimeOut"/> is configured.
        ''' <para></para>
        ''' This method is part of a retry loop that uses <see cref="NativeMethods.EnumThreadWindows"/>
        ''' to locate a modal dialog window (a message box) created on the current thread.
        ''' <para></para>
        ''' If the window is not found within <see cref="DevMessageBox.EnumWindowMaxTries"/> number of attempts, the process stops.
        ''' </summary>
        <DebuggerStepThrough>
        Private Sub FindDialog()

            If Me.enumWindowCurrentTries = DevMessageBox.EnumWindowMaxTries Then
                Throw New Exception($"Can't find {NameOf(DevMessageBox)} dialog window. ({NameOf(DevMessageBox.FindDialog)} function attempts: {Me.enumWindowCurrentTries})")
            End If

            Static callback As New Delegates.EnumThreadWindowsProc(AddressOf Me.EnumThreadWindowsCallback)

            If NativeMethods.EnumThreadWindows(NativeMethods.GetCurrentThreadId(), callback, IntPtr.Zero) Then

                If Interlocked.Increment(Me.enumWindowCurrentTries) <= DevMessageBox.EnumWindowMaxTries Then
                    Thread.Sleep(5)
                    Me.ownerControl.BeginInvoke(New System.Windows.Forms.MethodInvoker(AddressOf Me.FindDialog))
                End If
            End If

            If Me.TimeOut.TotalMilliseconds > 0 Then

                Me.countdownSeconds = CInt(Me.TimeOut.TotalMilliseconds) \ 1000
                Me.CountdownTimer?.Dispose()
                Me.CountdownTimer = New System.Windows.Forms.Timer With {
                    .Interval = 1000,
                    .Enabled = True
                }
                Me.CountdownTimer.Start()

                Me.DestroyTimer?.Dispose()
                Me.DestroyTimer = New System.Windows.Forms.Timer With {
                    .Interval = CInt(Me.TimeOut.TotalMilliseconds),
                    .Enabled = True
                }
                Me.DestroyTimer.Start()
            End If
        End Sub

        ''' <summary>
        ''' Identifies whether a given window handle corresponds to a dialog window, and if so, centers it
        ''' relative to its owner window, ensuring proper visibility and position.
        ''' </summary>
        ''' 
        ''' <param name="hWnd">
        ''' The handle to the window being evaluated.
        ''' </param>
        ''' 
        ''' <param name="lParam">
        ''' Additional message-specific information (not used in this implementation) 
        ''' passed during enumeration in  <see cref="NativeMethods.EnumThreadWindows"/> function.
        ''' </param>
        ''' 
        ''' <returns>
        ''' <see langword="False"/> if the window is identified as a dialog and repositioned (to stop enumerating windows);
        ''' otherwise, <see langword="True"/> to continue enumerating windows.
        ''' </returns>
        <DebuggerStepThrough>
        Private Function EnumThreadWindowsCallback(hWnd As IntPtr, lParam As IntPtr) As Boolean

            ' Check if <hWnd> is a dialog window.
            Dim sb As New StringBuilder(capacity:=8, maxCapacity:=8)
            Dim result As Integer = NativeMethods.GetClassName(hWnd, sb, sb.Capacity)
            If sb.ToString() <> "#32770" Then
                Return True
            End If

            ' Get the control that displays the text.
            Dim hText As IntPtr = NativeMethods.GetDlgItem(hWnd, &HFFFFI)
            Me.DialogWindowHandle = hWnd

            ' Get the owner rectangle.
            Dim ownerRect As Rectangle
            If Me.ownerWindow_ IsNot Nothing Then
                DevMessageBox.SetBoundsIfWindowIsVisibleAndNotOffScreen(Me.ownerWindow_, ownerRect)
            End If
            ' The details box will be centered to the current screen bounds as fallback.
            If ownerRect = Rectangle.Empty Then
                ownerRect = Screen.FromPoint(Cursor.Position).Bounds
            End If

            ' Get the dialog rectangle.
            Dim dlgRect As NativeRectangle
            NativeMethods.GetWindowRect(hWnd, dlgRect)

            ' Center the dialog window to the owner window.
            NativeMethods.SetWindowPos(hWnd:=hWnd, hWndInsertAfter:=IntPtr.Zero,
                                       x:=ownerRect.Left + ((ownerRect.Width - dlgRect.Right + dlgRect.Left) \ 2I),
                                       y:=ownerRect.Top + ((ownerRect.Height - dlgRect.Bottom + dlgRect.Top) \ 2I),
                                       cx:=0, cy:=0,
                                       uFlags:=SetWindowPosFlags.IgnoreResize Or SetWindowPosFlags.ShowWindow)

            ' Ensure the dialog is visible in the screen, and activate the window.
            NativeMethods.SendMessage(hWnd, DialogBoxMessages.Reposition, IntPtr.Zero, IntPtr.Zero)
            NativeMethods.SetActiveWindow(hWnd)

            ' Add the countdown label to dialog window.
            Dim needsSubclassing As Boolean =
                (Me.TimeOut.TotalMilliseconds > 0 AndAlso Me.countdownLabelHandle = IntPtr.Zero) OrElse
                (Me.DialogBackgroundColor <> Color.Empty AndAlso Me.previousDlgWndProc = IntPtr.Zero)

            If needsSubclassing Then
                If Me.TimeOut.TotalMilliseconds > 0 AndAlso Me.countdownLabelHandle = IntPtr.Zero Then
                    Me.CreateCountdownLabelWindow(Me.DialogWindowHandle)
                End If

                ' Replace dialog window procedure (WndProc).
                Static newDlgWndProc As WndProc = AddressOf Me.CustomDlgWndProc
                Dim newDlgWndProcPtr As IntPtr = Marshal.GetFunctionPointerForDelegate(newDlgWndProc)
#Disable Warning BC40000' Type or member is obsolete
                Me.previousDlgWndProc = If(Environment.Is64BitProcess,
                    NativeMethods.SetWindowLongPtr(Me.DialogWindowHandle, WindowLongValues.WndProc, newDlgWndProcPtr),
                    New IntPtr(NativeMethods.SetWindowLong(Me.DialogWindowHandle, WindowLongValues.WndProc, CUInt(newDlgWndProcPtr))))
#Enable Warning BC40000' Type or member is obsolete

                ' Force a repaint so the custom WndProc processes WM_ERASEBKGND.
                If Me.DialogBackgroundColor <> Color.Empty Then
                    Dim clientRect As NativeRectangle
                    NativeMethods.GetClientRect(Me.DialogWindowHandle, clientRect)
                    NativeMethods.InvalidateRect(Me.DialogWindowHandle, clientRect, [erase]:=True)
                    NativeMethods.UpdateWindow(Me.DialogWindowHandle)
                End If
            End If

            Const UseImmersiveDarkMode As Integer = 20
            ' Apply dark title bar.
            If Me.UseDarkTitleBar Then
                Dim flagValue As Integer = 1
                NativeMethods.DwmSetWindowAttribute(
                    Me.DialogWindowHandle, UseImmersiveDarkMode,
                    flagValue, Marshal.SizeOf(GetType(Integer)))

                ' Force DWM to repaint the title bar by toggling the window size.
                Dim dlgWinRect As NativeRectangle
                NativeMethods.GetWindowRect(Me.DialogWindowHandle, dlgWinRect)
                Dim dlgW As Integer = dlgWinRect.Right - dlgWinRect.Left
                Dim dlgH As Integer = dlgWinRect.Bottom - dlgWinRect.Top
                NativeMethods.SetWindowPos(Me.DialogWindowHandle, IntPtr.Zero, 0, 0, dlgW, dlgH + 1,
                                           SetWindowPosFlags.IgnoreMove Or SetWindowPosFlags.IgnoreZOrder)
                NativeMethods.SetWindowPos(Me.DialogWindowHandle, IntPtr.Zero, 0, 0, dlgW, dlgH,
                                           SetWindowPosFlags.IgnoreMove Or SetWindowPosFlags.IgnoreZOrder)
            End If

            ' Stop the EnumThreadWndProc callback by returning False.
            Return False

        End Function

        ''' <summary>
        ''' Executes the specified message box function and handles setup for different application types
        ''' (Console or DLL). It ensures the message box is displayed correctly, particularly in console applications,
        ''' by centering the window on the console's rectangle using a transparent owner form.
        ''' </summary>
        ''' 
        ''' <param name="msgboxfunc">
        ''' A <see cref="Func(Of DialogResult)"/> delegate that encapsulates the message box invocation logic.
        ''' </param>
        ''' 
        ''' <returns>
        ''' A <see cref="DialogResult"/> representing the result of the message box interaction.
        ''' </returns>
        ''' 
        ''' <exception cref="NotSupportedException">
        ''' Thrown when the entry assembly is a DLL.
        ''' </exception>
        <DebuggerStepThrough>
        Protected Function ReturnMessageBoxResult(msgboxfunc As Func(Of DialogResult)) As DialogResult

            If Me.ownerControl Is Nothing OrElse
               (
                TypeOf Me.ownerControl Is TransparentControlsForm AndAlso
                Me.ownerControl.Name = $"{NameOf(DevMessageBox)}_{NameOf(TransparentControlsForm)}_Windows"
               ) Then

                Using ctrl As New Control()
                    Dim screenBounds As Rectangle = Screen.FromControl(ctrl).Bounds
                    Dim centerPoint As New Point(
                        screenBounds.Left + screenBounds.Width \ 2,
                        screenBounds.Top + screenBounds.Height \ 2
                    )
                    Me.transparentForm?.Dispose()
                    Me.transparentForm = New TransparentControlsForm(ctrl, centerPoint) With {
                        .Name = $"{NameOf(DevMessageBox)}_{NameOf(TransparentControlsForm)}_Windows"
                    }
                    Me.ownerControl = Me.transparentForm
                End Using
            End If

            Me.ownerControl.BeginInvoke(New System.Windows.Forms.MethodInvoker(AddressOf Me.FindDialog))
            Dim dlgResult As DialogResult = msgboxfunc.Invoke()

            If Me.TimeOut = TimeSpan.Zero Then
                Return dlgResult

            ElseIf Me.TimeOut.TotalMilliseconds > 0 AndAlso
                   (Me.CountdownTimer IsNot Nothing OrElse Me.DestroyTimer IsNot Nothing) Then
                Return dlgResult

            Else
                Return DialogResult.None

            End If

        End Function

        ''' <summary>
        ''' Custom window procedure (WndProc) that intercepts and processes window messages for a subclassed window.
        ''' <para></para>
        ''' Allows custom behavior to be implemented for specific messages before optionally delegating to the original procedure.
        ''' </summary>
        ''' 
        ''' <param name="hWnd">
        ''' A handle to the window receiving the message.
        ''' </param>
        ''' 
        ''' <param name="msg">
        ''' The identifier of the message being processed.
        ''' </param>
        ''' 
        ''' <param name="wParam">
        ''' Additional message-specific information.
        ''' </param>
        ''' 
        ''' <param name="lParam">
        ''' Additional message-specific information.
        ''' </param>
        ''' 
        ''' <returns>
        ''' The result of message processing. The return value depends on the message and may represent a pointer, a handle, 
        ''' or a message-specific value. If the message is not handled, the result of the original window procedure is returned.
        ''' </returns>
        <DebuggerStepThrough>
        Private Function CustomDlgWndProc(hWnd As IntPtr, msg As UInteger, wParam As IntPtr, lParam As IntPtr) As IntPtr

            Const WM_CtlColorDlg As Integer = &H136
            Const WM_CtlColorBtn As Integer = &H135
            Const WM_CtlColorStatic As Integer = &H138
            Const WM_EraseBkgnd As Integer = &H14
            Const WM_Paint As Integer = &HF

            Select Case msg

                Case WM_CtlColorStatic
                    Dim hdc As IntPtr = wParam

                    ' Countdown label gets its own dedicated colors.
                    If lParam = Me.countdownLabelHandle Then
                        NativeMethods.SetTextColor(hdc, New NativeColor(Me.CountdownForegroundColor))
                        NativeMethods.SetBkColor(hdc, New NativeColor(Me.CountdownBackgroundColor))
                        NativeMethods.SetBkMode(hdc, BackgroundMode.Transparent)

                        If Me.countdownBrush = IntPtr.Zero Then
                            Me.countdownBrush = NativeMethods.CreateSolidBrush(ColorTranslator.ToWin32(Me.CountdownBackgroundColor))
                        End If
                        Return Me.countdownBrush
                    End If

                    ' All other static controls (text label, icon background) use dialog colors.
                    If Me.DialogBackgroundColor <> Color.Empty Then
                        If Me.DialogForegroundColor <> Color.Empty Then
                            NativeMethods.SetTextColor(hdc, New NativeColor(Me.DialogForegroundColor))
                        End If
                        NativeMethods.SetBkMode(hdc, BackgroundMode.Transparent)

                        If Me.dialogBackgroundBrush = IntPtr.Zero Then
                            Me.dialogBackgroundBrush = NativeMethods.CreateSolidBrush(ColorTranslator.ToWin32(Me.DialogBackgroundColor))
                        End If
                        Return Me.dialogBackgroundBrush
                    End If

                Case WM_CtlColorDlg
                    ' Paints the dialog surface itself.
                    If Me.DialogBackgroundColor <> Color.Empty Then
                        If Me.dialogBackgroundBrush = IntPtr.Zero Then
                            Me.dialogBackgroundBrush = NativeMethods.CreateSolidBrush(ColorTranslator.ToWin32(Me.DialogBackgroundColor))
                        End If
                        Return Me.dialogBackgroundBrush
                    End If

                Case WM_CtlColorBtn
                    ' Paints the area behind themed buttons.
                    If Me.DialogBackgroundColor <> Color.Empty Then
                        NativeMethods.SetBkMode(wParam, BackgroundMode.Transparent)

                        If Me.dialogBackgroundBrush = IntPtr.Zero Then
                            Me.dialogBackgroundBrush = NativeMethods.CreateSolidBrush(ColorTranslator.ToWin32(Me.DialogBackgroundColor))
                        End If
                        Return Me.dialogBackgroundBrush
                    End If

                Case WM_EraseBkgnd
                    If Me.DialogBackgroundColor <> Color.Empty Then
                        Dim clientRect As NativeRectangle
                        NativeMethods.GetClientRect(hWnd, clientRect)

                        If Me.dialogBackgroundBrush = IntPtr.Zero Then
                            Me.dialogBackgroundBrush = NativeMethods.CreateSolidBrush(ColorTranslator.ToWin32(Me.DialogBackgroundColor))
                        End If

                        NativeMethods.FillRect(wParam, clientRect, Me.dialogBackgroundBrush)
                        Return New IntPtr(1)
                    End If

                Case WM_Paint
                    Dim hasBackColor As Boolean = Me.DialogBackgroundColor <> Color.Empty
                    Dim hasFooterColor As Boolean = Me.DialogFooterBackgroundColor <> Color.Empty

                    If hasBackColor OrElse hasFooterColor Then
                        Dim paintResult As IntPtr = NativeMethods.CallWindowProc(
                            Me.previousDlgWndProc, hWnd, msg, wParam, lParam)

                        Dim hdc As IntPtr = NativeMethods.GetDCEx(
                            hWnd, IntPtr.Zero, GetDCExFlags.Cache Or GetDCExFlags.ClipChildren)

                        If hdc <> IntPtr.Zero Then
                            Dim clientRect As NativeRectangle
                            NativeMethods.GetClientRect(hWnd, clientRect)

                            ' Paint entire client area with background color.
                            If hasBackColor Then
                                If Me.dialogBackgroundBrush = IntPtr.Zero Then
                                    Me.dialogBackgroundBrush = NativeMethods.CreateSolidBrush(
                                        ColorTranslator.ToWin32(Me.DialogBackgroundColor))
                                End If
                                NativeMethods.FillRect(hdc, clientRect, Me.dialogBackgroundBrush)
                            End If

                            ' Paint footer region on top.
                            If hasFooterColor Then
                                Dim hBtn As IntPtr = IntPtr.Zero
                                For btnId As Integer = 1 To 9
                                    hBtn = NativeMethods.GetDlgItem(hWnd, btnId)
                                    If hBtn <> IntPtr.Zero Then
                                        Exit For
                                    End If
                                Next

                                If hBtn <> IntPtr.Zero Then
                                    Dim btnScreenRect As NativeRectangle
                                    NativeMethods.GetWindowRect(hBtn, btnScreenRect)

                                    Dim dlgScreenRect As NativeRectangle
                                    NativeMethods.GetWindowRect(hWnd, dlgScreenRect)

                                    ' Client area bottom edge ≈ dialog bottom edge (DWM border is negligible).
                                    ' So client top in screen coords = dialog bottom - client height.
                                    Dim clientScreenTop As Integer = dlgScreenRect.Bottom - clientRect.Bottom
                                    Dim footerTop As Integer = btnScreenRect.Top - clientScreenTop

                                    If footerTop > 0 AndAlso footerTop < clientRect.Bottom Then
                                        If Me.dialogFooterBackgroundBrush = IntPtr.Zero Then
                                            Me.dialogFooterBackgroundBrush = NativeMethods.CreateSolidBrush(
                                                ColorTranslator.ToWin32(Me.DialogFooterBackgroundColor))
                                        End If
                                        Dim footerRect As New NativeRectangle With {
                                            .Left = clientRect.Left,
                                            .Top = footerTop + (footerTop \ 2),
                                            .Right = clientRect.Right,
                                            .Bottom = clientRect.Bottom + footerTop
                                        }
                                        NativeMethods.FillRect(hdc, footerRect, Me.dialogFooterBackgroundBrush)
                                    End If
                                End If
                            End If

                            NativeMethods.ReleaseDC(hWnd, hdc)
                        End If

                        Return paintResult
                    End If

            End Select

            Return NativeMethods.CallWindowProc(Me.previousDlgWndProc, hWnd, msg, wParam, lParam)
        End Function

        ''' <summary>
        ''' Creates and positions the countdown label window inside the dialog window.
        ''' <para></para>
        ''' It uses native Win32 API calls to create and size the label dynamically based on the expected text size.
        ''' </summary>
        <DebuggerStepThrough>
        Protected Overridable Sub CreateCountdownLabelWindow(dlgHwnd As IntPtr)

            Dim dlgClientRect As NativeRectangle
            NativeMethods.GetClientRect(dlgHwnd, dlgClientRect)
            Dim dlgWidth As Integer = dlgClientRect.Right - dlgClientRect.Left
            Const minWidth As Integer = 340

            Dim labelWidth As Integer = 100
            Dim labelHeight As Integer = 20
            Dim labelX As Integer = 0
            Dim labelY As Integer = If(dlgWidth < minWidth, 0, dlgClientRect.Bottom - labelHeight)

            Dim formatted As String = $"{Math.Floor(Me.TimeOut.TotalHours):00}:{Me.TimeOut.Minutes:00}:{Me.TimeOut.Seconds:00}"
            Me.countdownLabelHandle = NativeMethods.CreateWindowEx(
                Me.CountdownWindowStyles,
                "STATIC",
                $" {formatted} ",
                WindowStyles.Child Or WindowStyles.Visible,
                labelX, labelY,
                labelWidth, labelHeight,
                dlgHwnd, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero
            )

            ' Measures the proper label's widow width and height.
            Const textToDraw As String = " 99:59:59 "
            Dim hdc As IntPtr = NativeMethods.GetDC(Me.countdownLabelHandle)
            Dim textSize As NativeSize
            NativeMethods.GetTextExtentPoint32(hdc, textToDraw, textToDraw.Length, textSize)
            NativeMethods.ReleaseDC(Me.countdownLabelHandle, hdc)
            labelWidth = textSize.Width
            labelHeight = textSize.Height
            labelY = If(dlgWidth < minWidth, 0, dlgClientRect.Bottom - labelHeight)

            NativeMethods.SetWindowPos(
                Me.countdownLabelHandle,
                IntPtr.Zero, labelX, labelY,
                labelWidth, labelHeight,
                SetWindowPosFlags.IgnoreZOrder Or SetWindowPosFlags.DontActivate Or SetWindowPosFlags.ShowWindow
            )

        End Sub

        ''' <summary>
        ''' If the provided window is within its screen working area, 
        ''' sets the bounds of the specified rectangle to the window's bounds; 
        ''' otherwise, leaves the rectangle unchanged.
        ''' </summary>
        ''' 
        ''' <param name="window">
        ''' The window whose bounds are evaluated.
        ''' </param>
        ''' 
        ''' <param name="refRect">
        ''' When the method returns, contains the window bounds if the window is within its screen working area; 
        ''' otherwise, the rectangle remains unchanged.
        ''' </param>
        <DebuggerStepThrough>
        Private Shared Sub SetBoundsIfWindowIsVisibleAndNotOffScreen(window As IWin32Window, ByRef refRect As Rectangle)

            If window Is Nothing OrElse window.Handle = IntPtr.Zero Then
                Return
            End If

            If Not NativeMethods.IsWindowVisible(window.Handle) Then
                Return
            End If

            Dim screen As Screen = Screen.FromHandle(window.Handle)

            Dim nativeRect As NativeRectangle
            If Not NativeMethods.GetWindowRect(window.Handle, nativeRect) Then
                Return
            End If

            Dim rect As Rectangle = nativeRect
            If screen.WorkingArea.IntersectsWith(rect) Then
                refRect = nativeRect
            End If
        End Sub

        ''' <summary>
        ''' Cleans-up managed and native resources used by this class.
        ''' <para></para>
        ''' This method must be called always before calling <see cref="DevMessageBox.ReturnMessageBoxResult"/> function.
        ''' </summary>
        <DebuggerStepThrough>
        Protected Overridable Sub CleanUp()

            If Me.CountdownTimer IsNot Nothing Then
                Me.CountdownTimer.Stop()
                Me.CountdownTimer.Dispose()
                Me.CountdownTimer = Nothing
            End If

            If Me.DestroyTimer IsNot Nothing Then
                Me.DestroyTimer.Enabled = False
                Me.DestroyTimer.Dispose()
                Me.DestroyTimer = Nothing
            End If

            If Me.transparentForm IsNot Nothing Then
                Me.transparentForm.Dispose()
                Me.transparentForm = Nothing
            End If

            If Me.dialogBackgroundBrush <> IntPtr.Zero Then
                NativeMethods.DeleteObject(Me.dialogBackgroundBrush)
                Me.dialogBackgroundBrush = IntPtr.Zero
            End If

            If Me.dialogFooterBackgroundBrush <> IntPtr.Zero Then
                NativeMethods.DeleteObject(Me.dialogFooterBackgroundBrush)
                Me.dialogFooterBackgroundBrush = IntPtr.Zero
            End If

            If Me.countdownBrush <> IntPtr.Zero Then
                NativeMethods.DeleteObject(Me.countdownBrush)
                Me.countdownBrush = IntPtr.Zero
            End If

            If Me.countdownLabelHandle <> IntPtr.Zero Then
                NativeMethods.DestroyWindow(Me.countdownLabelHandle)
                Me.countdownLabelHandle = IntPtr.Zero
            End If

            If Me.DialogWindowHandle <> IntPtr.Zero Then
                NativeMethods.EndDialog(Me.DialogWindowHandle, IntPtr.Zero) ' DialogResult.None
                Me.DialogWindowHandle = IntPtr.Zero
            End If

        End Sub

#End Region

#Region " Event Handlers "

        ''' <summary>
        ''' Handles the <see cref="System.Windows.Forms.Timer.Tick"/> event of the <see cref="DevMessageBox.CountdownTimer"/> control.
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
        Protected Overridable Sub CountdownTimer_Tick(sender As Object, e As EventArgs) Handles CountdownTimer.Tick

            If Me.countdownLabelHandle <> IntPtr.Zero AndAlso Me.countdownSeconds <> 0 Then
                Dim ts As TimeSpan = TimeSpan.FromSeconds(Interlocked.Decrement(Me.countdownSeconds))
                Dim formatted As String = $"{Math.Floor(ts.TotalHours):00}:{ts.Minutes:00}:{ts.Seconds:00}"
                NativeMethods.SetWindowText(Me.countdownLabelHandle, $" {formatted} ")
            End If
        End Sub

        ''' <summary>
        ''' Handles the <see cref="System.Windows.Forms.Timer.Tick"/> event of the <see cref="DevMessageBox.DestroyTimer"/> control.
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
        Protected Overridable Sub DestroyTimer_Tick(sender As Object, e As EventArgs) Handles DestroyTimer.Tick

            Me.CleanUp()
        End Sub

#End Region

#Region " IDisposable Implementation "

        ''' <summary>
        ''' Flag to detect redundant calls when disposing.
        ''' </summary>
        Private isDisposed As Boolean = False

        ''' <summary>
        ''' Releases all the resources used by this instance.
        ''' </summary>
        <DebuggerStepThrough>
        Public Sub Dispose() Implements IDisposable.Dispose
            Me.Dispose(isDisposing:=True)
            GC.SuppressFinalize(Me)
        End Sub

        ''' <summary>
        ''' Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        ''' Releases unmanaged and, optionally, managed resources.
        ''' </summary>
        '''
        ''' <param name="isDisposing">
        ''' <see langword="True"/>  to release both managed and unmanaged resources; 
        ''' <see langword="False"/> to release only unmanaged resources.
        ''' </param>
        <DebuggerStepThrough>
        Protected Overridable Sub Dispose(isDisposing As Boolean)

            If (Not Me.isDisposed) AndAlso isDisposing Then

                Me.CleanUp()
                Me.enumWindowCurrentTries = 0
                Me.countdownSeconds = 0
                Me.consoleRect = Rectangle.Empty
            End If

            Me.isDisposed = True
        End Sub

#End Region

    End Class

    ' ---------------------------------------------------------------------------------------------------------------------
    ' PREVIOUS IMPLEMENTATION.
    ' ------------------------
    '
    ' This implementation should behave identically to the current one, 
    ' except that it does not support customization of the background and footer appearance.
    '
    ' The logic for centering the dialog relative to its owner window has been simplified and corrected
    ' in the current implementation (not this one), eliminating the need to use System.Reflection.
    ' Refer to the 'EnumThreadWindowsCallback' and 'ReturnMessageBoxResult' functions to keep track of the related changes.
    ' ---------------------------------------------------------------------------------------------------------------------
    '
    '    ''' <summary>
    '    ''' Represents a <see cref="MessageBox"/> that will be displayed centered to the 
    '    ''' specified <see cref="Form"/> or <see cref="Control"/>.
    '    ''' </summary>
    '    '''
    '    ''' <example> This is a code example.
    '    ''' <code language="VB.NET">
    '    ''' ' Windows Forms Application
    '    ''' 
    '    ''' Using msg As New DevMessageBox(owner:=Me, TimeSpan.FromSeconds(5))
    '    ''' 
    '    '''     Dim dlgresult As DialogResult = msg.Show("Question", "Title", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
    '    '''     Console.WriteLine(dlgresult.ToString())
    '    ''' End Using
    '    ''' </code>
    '    ''' </example>
    '    ''' 
    '    ''' <example> This is a code example.
    '    ''' <code language="VB.NET">
    '    ''' ' Console Application
    '    ''' 
    '    ''' Dim ownerWindow As New ConsoleWindowWrapper(UtilConsole.Handle)
    '    ''' DevCase.Win32.NativeMethods.SetForegroundWindow(ownerWindow.Handle)
    '    ''' 
    '    ''' Using msg As New DevMessageBox(ownerWindow, TimeSpan.FromSeconds(5))
    '    ''' 
    '    '''     Dim dlgresult As DialogResult = msg.Show("Question", "Title", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
    '    '''     Console.WriteLine(dlgresult.ToString())
    '    ''' End Using
    '    ''' </code>
    '    ''' </example>
    '    <SupportedOSPlatform("windows")>
    '    Public Class DevMessageBox : Implements IDisposable
    '
    '#Region " Private Fields "
    '
    '        ''' <summary>
    '        ''' Keeps track of the current amount of tries to find this <see cref="DevMessageBox"/> dialog.
    '        ''' </summary>
    '        Private enumWindowCurrentTries As Integer
    '
    '        ''' <summary>
    '        ''' The maximum amount of tries to try find this <see cref="DevMessageBox"/> dialog.
    '        ''' </summary>
    '        Protected Const EnumWindowMaxTries As Integer = 100
    '
    '        ''' <summary>
    '        ''' Number of seconds remaining before the message box closes automatically.
    '        ''' </summary>
    '        Private countdownSeconds As Integer
    '
    '        ''' <summary>
    '        ''' Handle to the native countdown label window created via WinAPI.
    '        ''' <para></para>
    '        ''' Used to update the label text and manage its visibility during the timeout countdown.
    '        ''' </summary>
    '        Protected countdownLabelHandle As IntPtr
    '
    '        ''' <summary>
    '        ''' A <see cref="System.Windows.Forms.Timer"/> used to decrement the countdown and update the label each second.
    '        ''' </summary>
    '        Protected WithEvents CountdownTimer As System.Windows.Forms.Timer
    '
    '        ''' <summary>
    '        ''' A <see cref="System.Windows.Forms.Timer"/> that keeps track of <see cref="DevMessageBox.TimeOut"/> value to close this <see cref="DevMessageBox"/>.
    '        ''' </summary>
    '        Protected WithEvents DestroyTimer As System.Windows.Forms.Timer
    '
    '        ''' <summary>
    '        ''' The messagebox main window handle (hwnd).
    '        ''' </summary>
    '        Protected DialogWindowHandle As IntPtr
    '
    '        ''' <summary>
    '        ''' Flag to determine the entry assembly type.
    '        ''' </summary>
    '        Protected entryPeFileKind As PEFileKinds
    '
    '        ''' <summary>
    '        ''' Keeps track of the console size and position 
    '        ''' when <see cref="entryPeFileKind"/> is <see cref="PEFileKinds.ConsoleApplication"/>.
    '        ''' </summary>
    '        Private consoleRect As NativeRectangle
    '
    '        ''' <summary>
    '        ''' A transparent form that enables to enumerate thread windows to find the dialog window 
    '        ''' when <see cref="entryPeFileKind"/> is <see cref="PEFileKinds.ConsoleApplication"/>.
    '        ''' </summary>
    '        Private transparentForm As TransparentControlsForm
    '
    '        ''' <summary>
    '        ''' Stores a handle to the previous window procedure (WndProc) for the dialog window, 
    '        ''' used to call the original window procedure after subclassing.
    '        ''' </summary>
    '        Private previousDlgWndProc As IntPtr
    '
    '        ''' <summary>
    '        ''' Stores a handle to the native brush used to paint the countdown label background.
    '        ''' </summary>
    '        Private countdownBrush As IntPtr = IntPtr.Zero
    '
    '        ''' <summary>
    '        ''' The <see cref="Control"/> that owns this <see cref="DevMessageBox"/>.
    '        ''' <para></para>
    '        ''' This value is <c>null</c> if no <see cref="Control"/> has been passed to the <see cref="DevMessageBox"/> constructor.
    '        ''' </summary>
    '        Protected ownerControl As Control
    '
    '#End Region
    '
    '#Region " Properties "
    '
    '        ''' <summary>
    '        ''' Gets the <see cref="IWin32Window"/> that owns this <see cref="DevMessageBox"/>.
    '        ''' <para></para>
    '        ''' This value is <c>null</c> if no <see cref="IWin32Window"/> has been passed to the <see cref="DevMessageBox"/> constructor.
    '        ''' </summary>
    '        '''
    '        ''' <value>
    '        ''' The the <see cref="IWin32Window"/> that owns this <see cref="DevMessageBox"/>; 
    '        ''' or <c>null</c> if no <see cref="IWin32Window"/> has been passed to the <see cref="DevMessageBox"/> constructor.
    '        ''' </value>
    '        Public ReadOnly Property OwnerWindow As IWin32Window
    '            Get
    '                If Me.ownerWindow_ IsNot Nothing Then
    '                    Return Me.ownerWindow_
    '
    '                ElseIf Me.ownerControl IsNot Nothing Then
    '                    Return NativeWindow.FromHandle(Me.ownerControl.Handle)
    '
    '                Else
    '                    Return Nothing
    '
    '                End If
    '            End Get
    '        End Property
    '        ''' <summary>
    '        ''' (Backing Field)
    '        ''' <para></para>
    '        ''' The <see cref="IWin32Window"/> that owns this <see cref="DevMessageBox"/>.
    '        ''' <para></para>
    '        ''' This value is <c>null</c> if no <see cref="IWin32Window"/> has been passed to the <see cref="DevMessageBox"/> constructor.
    '        ''' </summary>
    '        Private ReadOnly ownerWindow_ As IWin32Window
    '
    '        ''' <summary>
    '        ''' Gets or sets the time interval that this <see cref="DevMessageBox"/> will close automatically.
    '        ''' <para></para>
    '        ''' Minimum allowed value is 2000 total milliseconds (2 seconds).
    '        ''' <para></para>
    '        ''' Maximum allowed value is 359999000 total milliseconds (99h:59m:59s).
    '        ''' <para></para>
    '        ''' If this value is zero (<see cref="TimeSpan.Zero"/>) the <see cref="DevMessageBox"/> will not display a countdown.
    '        ''' <para></para>
    '        ''' Default value is <see cref="TimeSpan.Zero"/>.
    '        ''' </summary>
    '        Public Property TimeOut As TimeSpan
    '            Get
    '                Return Me.timeOut_
    '            End Get
    '            <DebuggerStepThrough>
    '            Set(value As TimeSpan)
    '                If value = Nothing Then
    '                    Me.timeOut_ = TimeSpan.Zero
    '                Else
    '                    Dim ms As Double = value.TotalMilliseconds
    '                    Const MinMs As Double = 2000      ' 2 seconds.
    '                    Const MaxMs As Double = 359999000 ' 99h:59m:59s.
    '
    '                    If ms < MinMs Then
    '                        Throw New ArgumentOutOfRangeException(
    '                            NameOf(DevMessageBox.TimeOut), value,
    '                            $"Timeout value is out of range, it must be at least 2 seconds or higher."
    '                        )
    '                    End If
    '
    '                    If ms > MaxMs Then
    '                        Throw New ArgumentOutOfRangeException(
    '                            NameOf(DevMessageBox.TimeOut), value,
    '                            $"Timeout value is out of range, total milliseconds can't exceed {MaxMs} (99h:59m:59s)."
    '                        )
    '                    End If
    '                End If
    '
    '                timeOut_ = value
    '            End Set
    '        End Property
    '
    '        ''' <summary>
    '        ''' (Backing Field)
    '        ''' <para></para>
    '        ''' The time interval that this <see cref="DevMessageBox"/> will close automatically.
    '        ''' </summary>
    '        Private timeOut_ As TimeSpan = TimeSpan.Zero
    '
    '        ''' <summary>
    '        ''' Gets or sets the <see cref="WindowStylesEx"/> used to create the countdown label.
    '        ''' </summary>
    '        Public Property CountdownWindowStyles As WindowStylesEx = WindowStylesEx.Default
    '
    '        ''' <summary>
    '        ''' Gets or sets the color used to paint the countdown label foreground.
    '        ''' </summary>
    '        Public Property CountdownForegroundColor As Color = SystemColors.ControlText
    '
    '        ''' <summary>
    '        ''' Gets or sets the color used to paint the countdown label background.
    '        ''' </summary>
    '        Public Property CountdownBackgroundColor As Color = SystemColors.Control
    '
    '#End Region
    '
    '#Region " Constructors "
    '
    '        ''' <summary>
    '        ''' Initializes a new instance of the <see cref="DevMessageBox"/> class.
    '        ''' </summary>
    '        '''
    '        ''' <param name="owner">
    '        ''' The <see cref="IWin32Window"/> that acts as the owner of this <see cref="DevMessageBox"/>.
    '        ''' <para></para>
    '        ''' Assigning an owner enables modal behavior, blocking input until this message box closes,
    '        ''' and ensures the message box is centered relative to the owner's bounds.
    '        ''' <para></para>
    '        ''' This value can be <c>null</c>, in such case the message box will behave as a standard modal dialog window 
    '        ''' centered to the current screen bounds.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="timeOut">
    '        ''' The time interval that this <see cref="DevMessageBox"/> will close automatically.
    '        ''' <para></para>
    '        ''' Minimum allowed value is 2000 total milliseconds (2 seconds).
    '        ''' <para></para>
    '        ''' Maximum allowed value is 359999000 total milliseconds (99h:59m:59s).
    '        ''' <para></para>
    '        ''' If this value is zero (<see cref="TimeSpan.Zero"/>) the <see cref="DevMessageBox"/> will not display a countdown.
    '        ''' <para></para>
    '        ''' Default value is <see cref="TimeSpan.Zero"/>.
    '        ''' </param>
    '        <DebuggerStepThrough>
    '        Public Sub New(owner As IWin32Window, Optional timeOut As TimeSpan = Nothing)
    '
    '            Me.TimeOut = timeOut
    '            Me.ownerWindow_ = owner
    '        End Sub
    '
    '        ''' <summary>
    '        ''' Initializes a new instance of the <see cref="DevMessageBox"/> class.
    '        ''' </summary>
    '        '''
    '        ''' <param name="owner">
    '        ''' The <see cref="Form"/> that acts as the owner of this <see cref="DevMessageBox"/>.
    '        ''' <para></para>
    '        ''' Assigning an owner enables modal behavior, blocking input until this message box closes,
    '        ''' and ensures the message box is centered relative to the owner's bounds.
    '        ''' <para></para>
    '        ''' This value can be <c>null</c>, in such case the message box will behave as a standard modal dialog window 
    '        ''' centered to the current screen bounds.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="timeOut">
    '        ''' The time interval that this <see cref="DevMessageBox"/> will close automatically.
    '        ''' <para></para>
    '        ''' Minimum allowed value is 2000 total milliseconds (2 seconds).
    '        ''' <para></para>
    '        ''' Maximum allowed value is 359999000 total milliseconds (99h:59m:59s).
    '        ''' <para></para>
    '        ''' If this value is zero (<see cref="TimeSpan.Zero"/>) the <see cref="DevMessageBox"/> will not display a countdown.
    '        ''' <para></para>
    '        ''' Default value is <see cref="TimeSpan.Zero"/>.
    '        ''' </param>
    '        <DebuggerStepThrough>
    '        Public Sub New(owner As Form, Optional timeOut As TimeSpan = Nothing)
    '
    '            Me.New(If(owner Is Nothing, Nothing, NativeWindow.FromHandle(owner.Handle)), timeOut)
    '            Me.ownerControl = owner
    '        End Sub
    '
    '        ''' <summary>
    '        ''' Initializes a new instance of the <see cref="DevMessageBox"/> class.
    '        ''' </summary>
    '        '''
    '        ''' <param name="owner">
    '        ''' The <see cref="Control"/> that acts as the owner of this <see cref="DevMessageBox"/>.
    '        ''' <para></para>
    '        ''' Assigning an owner enables modal behavior, blocking input until this message box closes,
    '        ''' and ensures the message box is centered relative to the owner's bounds.
    '        ''' <para></para>
    '        ''' This value can be <c>null</c>, in such case the message box will behave as a standard modal dialog window 
    '        ''' centered to the current screen bounds.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="timeOut">
    '        ''' The time interval that this <see cref="DevMessageBox"/> will close automatically.
    '        ''' <para></para>
    '        ''' Minimum allowed value is 2000 total milliseconds (2 seconds).
    '        ''' <para></para>
    '        ''' Maximum allowed value is 359999000 total milliseconds (99h:59m:59s).
    '        ''' <para></para>
    '        ''' If this value is zero (<see cref="TimeSpan.Zero"/>) the <see cref="DevMessageBox"/> will not display a countdown.
    '        ''' <para></para>
    '        ''' Default value is <see cref="TimeSpan.Zero"/>.
    '        ''' </param>
    '        <DebuggerStepThrough>
    '        Public Sub New(owner As Control, Optional timeOut As TimeSpan = Nothing)
    '
    '            Me.New(If(owner Is Nothing, Nothing, NativeWindow.FromHandle(owner.Handle)), timeOut)
    '            Me.ownerControl = owner
    '        End Sub
    '
    '        ''' <summary>
    '        ''' Initializes a new instance of the <see cref="DevMessageBox"/> class.
    '        ''' </summary>
    '        '''
    '        ''' <param name="owner">
    '        ''' The <see cref="ConsoleWindowWrapper"/> that acts as the owner of this <see cref="DevMessageBox"/>.
    '        ''' <para></para>
    '        ''' Assigning an owner enables modal behavior, blocking input until this message box closes,
    '        ''' and ensures the message box is centered relative to the owner's bounds.
    '        ''' <para></para>
    '        ''' This value can be <c>null</c>, in such case the message box will behave as a standard modal dialog window 
    '        ''' centered to the current screen bounds.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="timeOut">
    '        ''' The time interval that this <see cref="DevMessageBox"/> will close automatically.
    '        ''' <para></para>
    '        ''' Minimum allowed value is 2000 total milliseconds (2 seconds).
    '        ''' <para></para>
    '        ''' Maximum allowed value is 359999000 total milliseconds (99h:59m:59s).
    '        ''' <para></para>
    '        ''' If this value is zero (<see cref="TimeSpan.Zero"/>) the <see cref="DevMessageBox"/> will not display a countdown.
    '        ''' <para></para>
    '        ''' Default value is <see cref="TimeSpan.Zero"/>.
    '        ''' </param>
    '        <DebuggerStepThrough>
    '        Public Sub New(owner As ConsoleWindowWrapper, Optional timeOut As TimeSpan = Nothing)
    '
    '            Me.New(DirectCast(owner, IWin32Window), timeOut)
    '        End Sub
    '
    '        ''' <summary>
    '        ''' Initializes a new instance of the <see cref="DevMessageBox"/> class 
    '        ''' without assigning an owner window.
    '        ''' </summary>
    '        <DebuggerStepThrough>
    '        Public Sub New()
    '
    '            Me.New(DirectCast(Nothing, IWin32Window), TimeSpan.Zero)
    '        End Sub
    '
    '#End Region
    '
    '#Region " Public Methods "
    '
    '        ''' <summary>
    '        ''' Displays a message box with specified text.
    '        ''' </summary>
    '        '''
    '        ''' <param name="text">
    '        ''' The text to display in the message box.
    '        ''' </param>
    '        '''
    '        ''' <returns>
    '        ''' One of the <see cref="DialogResult"/> values.
    '        ''' </returns>
    '        <DebuggerStepThrough>
    '        Public Overridable Function Show(text As String) As DialogResult
    '
    '            Me.CleanUp()
    '            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(Me.OwnerWindow, text))
    '        End Function
    '
    '        ''' <summary>
    '        ''' Displays a message box with specified text and caption.
    '        ''' </summary>
    '        '''
    '        ''' <param name="text">
    '        ''' The text to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="caption">
    '        ''' The text to display in the title bar of the message box.
    '        ''' </param>
    '        '''
    '        ''' <returns>
    '        ''' One of the <see cref="DialogResult"/> values.
    '        ''' </returns>
    '        <DebuggerStepThrough>
    '        Public Overridable Function Show(text As String, caption As String) As DialogResult
    '
    '            Me.CleanUp()
    '            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(Me.OwnerWindow, text, caption))
    '        End Function
    '
    '        ''' <summary>
    '        ''' Displays a message box with specified text, caption, and buttons.
    '        ''' </summary>
    '        '''
    '        ''' <param name="text">
    '        ''' The text to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="caption">
    '        ''' The text to display in the title bar of the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="buttons">
    '        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
    '        ''' </param>
    '        '''
    '        ''' <returns>
    '        ''' One of the <see cref="DialogResult"/> values.
    '        ''' <para></para>
    '        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
    '        ''' </returns>
    '        <DebuggerStepThrough>
    '        Public Overridable Function Show(text As String, caption As String, buttons As MessageBoxButtons) As DialogResult
    '
    '            Me.CleanUp()
    '            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(Me.OwnerWindow, text, caption, buttons))
    '        End Function
    '
    '        ''' <summary>
    '        ''' Displays a message box with specified text, caption, buttons, and icon.
    '        ''' </summary>
    '        '''
    '        ''' <param name="text">
    '        ''' The text to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="caption">
    '        ''' The text to display in the title bar of the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="buttons">
    '        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="icon">
    '        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
    '        ''' </param>
    '        '''
    '        ''' <returns>
    '        ''' One of the <see cref="DialogResult"/> values.
    '        ''' <para></para>
    '        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
    '        ''' </returns>
    '        <DebuggerStepThrough>
    '        Public Overridable Function Show(text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon) As DialogResult
    '
    '            Me.CleanUp()
    '            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(Me.OwnerWindow, text, caption, buttons, icon))
    '        End Function
    '
    '        ''' <summary>
    '        ''' Displays a message box with the specified text, caption, buttons, icon, and default button.
    '        ''' </summary>
    '        '''
    '        ''' <param name="text">
    '        ''' The text to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="caption">
    '        ''' The text to display in the title bar of the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="buttons">
    '        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="icon">
    '        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="defaultButton">
    '        ''' One of the <see cref="MessageBoxDefaultButton"/> values that specifies the default button for the message box.
    '        ''' </param>
    '        '''
    '        ''' <returns>
    '        ''' One of the <see cref="DialogResult"/> values.
    '        ''' <para></para>
    '        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
    '        ''' </returns>
    '        <DebuggerStepThrough>
    '        Public Overridable Function Show(text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon, defaultButton As MessageBoxDefaultButton) As DialogResult
    '
    '            Me.CleanUp()
    '            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(Me.OwnerWindow, text, caption, buttons, icon, defaultButton))
    '        End Function
    '
    '        ''' <summary>
    '        ''' Displays a message box with the specified text, caption, buttons, icon, default button, and options.
    '        ''' </summary>
    '        '''
    '        ''' <param name="text">
    '        ''' The text to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="caption">
    '        ''' The text to display in the title bar of the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="buttons">
    '        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="icon">
    '        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="defaultButton">
    '        ''' One of the <see cref="MessageBoxDefaultButton"/> values that specifies the default button for the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="options">
    '        ''' One of the <see cref="MessageBoxOptions"/> values that specifies which display and association options will be used for 
    '        ''' the message box. 
    '        ''' <para></para>
    '        ''' You may pass in 0 if you wish to use the defaults.
    '        ''' </param>
    '        '''
    '        ''' <returns>
    '        ''' One of the <see cref="DialogResult"/> values.
    '        ''' <para></para>
    '        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
    '        ''' </returns>
    '        <DebuggerStepThrough>
    '        Public Overridable Function Show(text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon, defaultButton As MessageBoxDefaultButton, options As MessageBoxOptions) As DialogResult
    '
    '            Me.CleanUp()
    '            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(Me.OwnerWindow, text, caption, buttons, icon, defaultButton, options))
    '        End Function
    '
    '        ''' <summary>
    '        ''' Displays a message box with the specified text, caption, buttons, icon, default button, options, and Help button, 
    '        ''' using the specified Help file.
    '        ''' </summary>
    '        '''
    '        ''' <param name="text">
    '        ''' The text to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="caption">
    '        ''' The text to display in the title bar of the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="buttons">
    '        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="icon">
    '        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="defaultButton">
    '        ''' One of the <see cref="MessageBoxDefaultButton"/> values that specifies the default button for the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="options">
    '        ''' One of the <see cref="MessageBoxOptions"/> values that specifies which display and association options will be used for 
    '        ''' the message box. 
    '        ''' <para></para>
    '        ''' You may pass in 0 if you wish to use the defaults.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="helpFilePath">
    '        ''' The path and name of the Help file to display when the user clicks the Help button.
    '        ''' </param>
    '        '''
    '        ''' <returns>
    '        ''' One of the <see cref="DialogResult"/> values.
    '        ''' <para></para>
    '        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
    '        ''' </returns>
    '        <DebuggerStepThrough>
    '        Public Overridable Function Show(text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon, defaultButton As MessageBoxDefaultButton, options As MessageBoxOptions, helpFilePath As String) As DialogResult
    '
    '            Me.CleanUp()
    '            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(Me.OwnerWindow, text, caption, buttons, icon, defaultButton, options, helpFilePath))
    '        End Function
    '
    '        ''' <summary>
    '        ''' Displays a message box with the specified text, caption, buttons, icon, default button, options, and Help button.
    '        ''' </summary>
    '        '''
    '        ''' <param name="text">
    '        ''' The text to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="caption">
    '        ''' The text to display in the title bar of the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="buttons">
    '        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="icon">
    '        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="defaultButton">
    '        ''' One of the <see cref="MessageBoxDefaultButton"/> values that specifies the default button for the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="options">
    '        ''' One of the <see cref="MessageBoxOptions"/> values that specifies which display and association options will be used for 
    '        ''' the message box. 
    '        ''' <para></para>
    '        ''' You may pass in 0 if you wish to use the defaults.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="displayHelpButton">
    '        ''' <see langword="True"/> to show the Help button; otherwise, false. The default is <see langword="False"/>.
    '        ''' </param>
    '        '''
    '        ''' <returns>
    '        ''' One of the <see cref="DialogResult"/> values.
    '        ''' <para></para>
    '        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
    '        ''' </returns>
    '        <DebuggerStepThrough>
    '        Public Overridable Function Show(text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon, defaultButton As MessageBoxDefaultButton, options As MessageBoxOptions, displayHelpButton As Boolean) As DialogResult
    '
    '            Me.CleanUp()
    '            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(text, caption, buttons, icon, defaultButton, options, displayHelpButton))
    '        End Function
    '
    '        ''' <summary>
    '        ''' Displays a message box with the specified text, caption, buttons, icon, default button, options, and Help button, 
    '        ''' using the specified Help file and HelpNavigator.
    '        ''' </summary>
    '        '''
    '        ''' <param name="text">
    '        ''' The text to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="caption">
    '        ''' The text to display in the title bar of the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="buttons">
    '        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="icon">
    '        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="defaultButton">
    '        ''' One of the <see cref="MessageBoxDefaultButton"/> values that specifies the default button for the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="options">
    '        ''' One of the <see cref="MessageBoxOptions"/> values that specifies which display and association options will be used for 
    '        ''' the message box. 
    '        ''' <para></para>
    '        ''' You may pass in 0 if you wish to use the defaults.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="helpFilePath">
    '        ''' The path and name of the Help file to display when the user clicks the Help button.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="navigator">
    '        ''' One of the <see cref="HelpNavigator"/> values.
    '        ''' </param>
    '        '''
    '        ''' <returns>
    '        ''' One of the <see cref="DialogResult"/> values.
    '        ''' <para></para>
    '        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
    '        ''' </returns>
    '        <DebuggerStepThrough>
    '        Public Overridable Function Show(text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon, defaultButton As MessageBoxDefaultButton, options As MessageBoxOptions, helpFilePath As String, navigator As HelpNavigator) As DialogResult
    '
    '            Me.CleanUp()
    '            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(Me.OwnerWindow, text, caption, buttons, icon, defaultButton, options, helpFilePath, navigator))
    '        End Function
    '
    '        ''' <summary>
    '        ''' Displays a message box with the specified text, caption, buttons, icon, default button, options, and Help button, 
    '        ''' using the specified Help file and Help keyword.
    '        ''' </summary>
    '        '''
    '        ''' <param name="text">
    '        ''' The text to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="caption">
    '        ''' The text to display in the title bar of the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="buttons">
    '        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="icon">
    '        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="defaultButton">
    '        ''' One of the <see cref="MessageBoxDefaultButton"/> values that specifies the default button for the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="options">
    '        ''' One of the <see cref="MessageBoxOptions"/> values that specifies which display and association options will be used for 
    '        ''' the message box. 
    '        ''' <para></para>
    '        ''' You may pass in 0 if you wish to use the defaults.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="helpFilePath">
    '        ''' The path and name of the Help file to display when the user clicks the Help button.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="keyword">
    '        ''' The Help keyword to display when the user clicks the Help button.
    '        ''' </param>
    '        '''
    '        ''' <returns>
    '        ''' One of the <see cref="DialogResult"/> values.
    '        ''' <para></para>
    '        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
    '        ''' </returns>
    '        <DebuggerStepThrough>
    '        Public Overridable Function Show(text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon, defaultButton As MessageBoxDefaultButton, options As MessageBoxOptions, helpFilePath As String, keyword As String) As DialogResult
    '
    '            Me.CleanUp()
    '            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(Me.OwnerWindow, text, caption, buttons, icon, defaultButton, options, helpFilePath, keyword))
    '        End Function
    '
    '        ''' <summary>
    '        ''' Displays a message box with the specified text, caption, buttons, icon, default button, options, and Help button, 
    '        ''' using the specified Help file, HelpNavigator, and Help topic.
    '        ''' </summary>
    '        '''
    '        ''' <param name="text">
    '        ''' The text to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="caption">
    '        ''' The text to display in the title bar of the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="buttons">
    '        ''' One of the <see cref="MessageBoxButtons"/> values that specifies which buttons to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="icon">
    '        ''' One of the <see cref="MessageBoxIcon"/> values that specifies which icon  to display in the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="defaultButton">
    '        ''' One of the <see cref="MessageBoxDefaultButton"/> values that specifies the default button for the message box.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="options">
    '        ''' One of the <see cref="MessageBoxOptions"/> values that specifies which display and association options will be used for 
    '        ''' the message box. 
    '        ''' <para></para>
    '        ''' You may pass in 0 if you wish to use the defaults.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="helpFilePath">
    '        ''' The path and name of the Help file to display when the user clicks the Help button.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="navigator">
    '        ''' One of the <see cref="HelpNavigator"/> values.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="param">
    '        ''' The numeric ID of the Help topic to display when the user clicks the Help button.
    '        ''' </param>
    '        '''
    '        ''' <returns>
    '        ''' One of the <see cref="DialogResult"/> values.
    '        ''' <para></para>
    '        ''' If no button is pressed during the <see cref="DevMessageBox.TimeOut"/> period, it returns <see cref="DialogResult.None"/>.
    '        ''' </returns>
    '        <DebuggerStepThrough>
    '        Public Overridable Function Show(text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon, defaultButton As MessageBoxDefaultButton, options As MessageBoxOptions, helpFilePath As String, navigator As HelpNavigator, param As Object) As DialogResult
    '
    '            Me.CleanUp()
    '            Return Me.ReturnMessageBoxResult(Function() MessageBox.Show(Me.OwnerWindow, text, caption, buttons, icon, defaultButton, options, helpFilePath, navigator, param))
    '        End Function
    '
    '#End Region
    '
    '#Region " Private Methods "
    '
    '        ''' <summary>
    '        ''' Executes the specified message box function and handles setup for different application types
    '        ''' (Console or DLL). It ensures the message box is displayed correctly, particularly in console applications,
    '        ''' by centering the window on the console's rectangle using a transparent owner form.
    '        ''' </summary>
    '        ''' 
    '        ''' <param name="msgboxfunc">
    '        ''' A <see cref="Func(Of DialogResult)"/> delegate that encapsulates the message box invocation logic.
    '        ''' </param>
    '        ''' 
    '        ''' <returns>
    '        ''' A <see cref="DialogResult"/> representing the result of the message box interaction.
    '        ''' </returns>
    '        ''' 
    '        ''' <exception cref="NotSupportedException">
    '        ''' Thrown when the entry assembly is a DLL.
    '        ''' </exception>
    '        <DebuggerStepThrough>
    '        Protected Function ReturnMessageBoxResult(msgboxfunc As Func(Of DialogResult)) As DialogResult
    '
    '            Me.entryPeFileKind = UtilPortableExecutable.GetPEFileKind(Assembly.GetEntryAssembly())
    '
    '            Select Case entryPeFileKind
    '
    '                Case PEFileKinds.Dll
    '                    Throw New NotSupportedException($"Using {NameOf(DevMessageBox)} from a DLL call is not supported.")
    '
    '                Case PEFileKinds.ConsoleApplication
    '                    Using ctrl As New Control()
    '                        Dim consoleHandle As IntPtr = UtilConsole.Handle
    '                        Dim isWindowVisible As Boolean = NativeMethods.IsWindowVisible(consoleHandle)
    '
    '                        Dim screenBounds As Rectangle = Screen.FromHandle(consoleHandle).Bounds
    '                        If isWindowVisible Then
    '                            ' Tries to retrieve the console rectangle where to center the messagebox window.
    '                            NativeMethods.GetWindowRect(consoleHandle, Me.consoleRect)
    '                            Dim isOffScreen As Boolean = Not CType(Me.consoleRect, Rectangle).IntersectsWith(screenBounds)
    '                            If isOffScreen Then
    '                                Me.consoleRect = screenBounds
    '                            End If
    '                        Else
    '                            Me.consoleRect = screenBounds
    '                        End If
    '
    '                        Me.transparentForm?.Dispose()
    '                        Me.transparentForm = New TransparentControlsForm(ctrl, Nothing) With {
    '                            .Name = $"{NameOf(DevMessageBox)}_{NameOf(TransparentControlsForm)}_Console"
    '                        }
    '                        Me.ownerControl = Me.transparentForm
    '                    End Using
    '
    '                Case Else
    '                    If Me.ownerControl Is Nothing OrElse
    '                        (TypeOf Me.ownerControl Is TransparentControlsForm AndAlso
    '                         Me.ownerControl.Name = $"{NameOf(DevMessageBox)}_{NameOf(TransparentControlsForm)}_Windows") Then
    '
    '                        Using ctrl As New Control()
    '                            Dim screenBounds As Rectangle = Screen.FromControl(ctrl).Bounds
    '                            Dim centerPoint As New Point(
    '                                screenBounds.Left + screenBounds.Width \ 2,
    '                                screenBounds.Top + screenBounds.Height \ 2
    '                            )
    '                            Me.transparentForm?.Dispose()
    '                            Me.transparentForm = New TransparentControlsForm(ctrl, centerPoint) With {
    '                                .Name = $"{NameOf(DevMessageBox)}_{NameOf(TransparentControlsForm)}_Windows"
    '                            }
    '                            Me.ownerControl = Me.transparentForm
    '                        End Using
    '                    End If
    '            End Select
    '
    '            Me.ownerControl.BeginInvoke(New System.Windows.Forms.MethodInvoker(AddressOf Me.FindDialog))
    '            Dim dlgResult As DialogResult = msgboxfunc.Invoke()
    '
    '            If Me.TimeOut = TimeSpan.Zero Then
    '                Return dlgResult
    '
    '            ElseIf Me.TimeOut.TotalMilliseconds > 0 AndAlso
    '                   (Me.CountdownTimer IsNot Nothing OrElse Me.DestroyTimer IsNot Nothing) Then
    '                Return dlgResult
    '
    '            Else
    '                Return DialogResult.None
    '
    '            End If
    '
    '        End Function
    '
    '        ''' <summary>
    '        ''' Attempts to locate the message box dialog window by enumerating the thread's windows,
    '        ''' and initiates a retry mechanism if the dialog is not found immediately.
    '        ''' <para></para>
    '        ''' Also starts countdown and destruction timers if <see cref="DevMessageBox.TimeOut"/> is configured.
    '        ''' <para></para>
    '        ''' This method is part of a retry loop that uses <see cref="NativeMethods.EnumThreadWindows"/>
    '        ''' to locate a modal dialog window (a message box) created on the current thread.
    '        ''' <para></para>
    '        ''' If the window is not found within <see cref="DevMessageBox.EnumWindowMaxTries"/> number of attempts, the process stops.
    '        ''' </summary>
    '        <DebuggerStepThrough>
    '        Private Sub FindDialog()
    '
    '            If Me.enumWindowCurrentTries = DevMessageBox.EnumWindowMaxTries Then
    '                Throw New Exception($"Can't find {NameOf(DevMessageBox)} dialog window. ({NameOf(DevMessageBox.FindDialog)} function attempts: {Me.enumWindowCurrentTries})")
    '            End If
    '
    '            Static callback As New Delegates.EnumThreadWindowsProc(AddressOf Me.EnumThreadWindowsCallback)
    '
    '            If NativeMethods.EnumThreadWindows(NativeMethods.GetCurrentThreadId(), callback, IntPtr.Zero) Then
    '
    '                If Interlocked.Increment(Me.enumWindowCurrentTries) <= DevMessageBox.EnumWindowMaxTries Then
    '                    Thread.Sleep(5)
    '                    Me.ownerControl.BeginInvoke(New System.Windows.Forms.MethodInvoker(AddressOf Me.FindDialog))
    '                End If
    '            End If
    '
    '            If Me.TimeOut.TotalMilliseconds > 0 Then
    '
    '                Me.countdownSeconds = CInt(Me.TimeOut.TotalMilliseconds) \ 1000
    '                Me.CountdownTimer?.Dispose()
    '                Me.CountdownTimer = New System.Windows.Forms.Timer With {
    '                    .Interval = 1000,
    '                    .Enabled = True
    '                }
    '                Me.CountdownTimer.Start()
    '
    '                Me.DestroyTimer?.Dispose()
    '                Me.DestroyTimer = New System.Windows.Forms.Timer With {
    '                    .Interval = CInt(Me.TimeOut.TotalMilliseconds),
    '                    .Enabled = True
    '                }
    '                Me.DestroyTimer.Start()
    '            End If
    '        End Sub
    '
    '        ''' <summary>
    '        ''' Identifies whether a given window handle corresponds to a dialog window, and if so, centers it
    '        ''' relative to its owner window, ensuring proper visibility and position.
    '        ''' </summary>
    '        ''' 
    '        ''' <param name="hWnd">
    '        ''' The handle to the window being evaluated.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="lParam">
    '        ''' Additional message-specific information (not used in this implementation) 
    '        ''' passed during enumeration in  <see cref="NativeMethods.EnumThreadWindows"/> function.
    '        ''' </param>
    '        ''' 
    '        ''' <returns>
    '        ''' <see langword="False"/> if the window is identified as a dialog and repositioned (to stop enumerating windows);
    '        ''' otherwise, <see langword="True"/> to continue enumerating windows.
    '        ''' </returns>
    '        <DebuggerStepThrough>
    '        Private Function EnumThreadWindowsCallback(hWnd As IntPtr, lParam As IntPtr) As Boolean
    '
    '            ' Check if <hWnd> is a dialog window.
    '            Dim sb As New StringBuilder(260)
    '            Dim result As Integer = NativeMethods.GetClassName(hWnd, sb, sb.Capacity)
    '            If sb.ToString() <> "#32770" Then
    '                Return True
    '            End If
    '
    '            ' Get the control that displays the text.
    '            Dim hText As IntPtr = NativeMethods.GetDlgItem(hWnd, &HFFFFI)
    '            Me.DialogWindowHandle = hWnd
    '
    '            ' Get the owner rectangle.
    '            Dim ownerRect As Rectangle
    '            If Me.entryPeFileKind = PEFileKinds.WindowApplication Then
    '                ownerRect = Me.ownerControl.Bounds
    '                Dim ownerHandle As IntPtr = Me.ownerControl.Handle
    '
    '                Dim isWindowVisible As Boolean = NativeMethods.IsWindowVisible(ownerHandle)
    '
    '                Dim screenBounds As Rectangle = Screen.FromHandle(ownerHandle).Bounds
    '                If isWindowVisible Then
    '                    ' Tries to retrieve the window rectangle where to center the messagebox window.
    '                    NativeMethods.GetWindowRect(ownerHandle, Me.consoleRect)
    '                    Dim isOffScreen As Boolean = Not CType(Me.consoleRect, Rectangle).IntersectsWith(screenBounds)
    '                    If isOffScreen Then
    '                        ownerRect = screenBounds
    '                    End If
    '                Else
    '                    ownerRect = screenBounds
    '                End If
    '
    '            ElseIf Me.entryPeFileKind = PEFileKinds.ConsoleApplication Then
    '                ownerRect = If(Me.consoleRect <> Rectangle.Empty,
    '                             New Rectangle(New Point(consoleRect.Left, consoleRect.Top), New Size(consoleRect.Right - consoleRect.Left, consoleRect.Bottom - consoleRect.Top)),
    '                             Screen.FromRectangle(Me.transparentForm.Bounds).Bounds)
    '
    '            End If
    '
    '            ' Get the dialog rectangle.
    '            Dim dlgRect As NativeRectangle
    '            NativeMethods.GetWindowRect(hWnd, dlgRect)
    '
    '            ' Center the dialog window in the owner window.
    '            NativeMethods.SetWindowPos(hWnd:=hWnd, hWndInsertAfter:=IntPtr.Zero,
    '                                       x:=ownerRect.Left + ((ownerRect.Width - dlgRect.Right + dlgRect.Left) \ 2I),
    '                                       y:=ownerRect.Top + ((ownerRect.Height - dlgRect.Bottom + dlgRect.Top) \ 2I),
    '                                       cx:=0, cy:=0,
    '                                       uFlags:=SetWindowPosFlags.IgnoreResize Or SetWindowPosFlags.ShowWindow)
    '
    '            ' Ensure the dialog is visible in the screen, and activate the window.
    '            NativeMethods.SendMessage(hWnd, DialogBoxMessages.Reposition, IntPtr.Zero, IntPtr.Zero)
    '            NativeMethods.SetActiveWindow(hWnd)
    '
    '            ' Add the countdown label to dialog window.
    '            If Me.TimeOut.TotalMilliseconds > 0 AndAlso Me.countdownLabelHandle = IntPtr.Zero Then
    '                Me.CreateCountdownLabelWindow(Me.DialogWindowHandle)
    '
    '                ' Replace dialog window procedure (WndProc).
    '                Static newDlgWndProc As WndProc = AddressOf Me.CustomDlgWndProc
    '                Dim newDlgWndProcPtr As IntPtr = Marshal.GetFunctionPointerForDelegate(newDlgWndProc)
    '#Disable Warning BC40000 ' Type or member is obsolete
    '                Me.previousDlgWndProc = If(Environment.Is64BitProcess,
    '                    NativeMethods.SetWindowLongPtr(Me.DialogWindowHandle, WindowLongValues.WndProc, newDlgWndProcPtr),
    '                    New IntPtr(NativeMethods.SetWindowLong(Me.DialogWindowHandle, WindowLongValues.WndProc, CUInt(newDlgWndProcPtr))))
    '#Enable Warning BC40000 ' Type or member is obsolete
    '            End If
    '
    '            ' Stop the EnumThreadWndProc callback by returning False.
    '            Return False
    '
    '        End Function
    '
    '        ''' <summary>
    '        ''' Custom window procedure (WndProc) that intercepts and processes window messages for a subclassed window.
    '        ''' <para></para>
    '        ''' Allows custom behavior to be implemented for specific messages before optionally delegating to the original procedure.
    '        ''' </summary>
    '        ''' 
    '        ''' <param name="hWnd">
    '        ''' A handle to the window receiving the message.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="msg">
    '        ''' The identifier of the message being processed.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="wParam">
    '        ''' Additional message-specific information.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="lParam">
    '        ''' Additional message-specific information.
    '        ''' </param>
    '        ''' 
    '        ''' <returns>
    '        ''' The result of message processing. The return value depends on the message and may represent a pointer, a handle, 
    '        ''' or a message-specific value. If the message is not handled, the result of the original window procedure is returned.
    '        ''' </returns>
    '        <DebuggerStepThrough>
    '        Private Function CustomDlgWndProc(hWnd As IntPtr, msg As UInteger, wParam As IntPtr, lParam As IntPtr) As IntPtr
    '
    '            If msg = WindowMessages.WM_CtlColorStatic Then
    '                If lParam = Me.countdownLabelHandle Then
    '                    Dim hdc As IntPtr = wParam
    '                    NativeMethods.SetTextColor(hdc, New NativeColor(Me.CountdownForegroundColor))
    '                    NativeMethods.SetBkColor(hdc, New NativeColor(Me.CountdownBackgroundColor))
    '                    NativeMethods.SetBkMode(hdc, BackgroundMode.Transparent)
    '
    '                    If Me.countdownBrush = IntPtr.Zero Then
    '                        Me.countdownBrush = NativeMethods.CreateSolidBrush(ColorTranslator.ToWin32(Me.CountdownBackgroundColor))
    '                    End If
    '
    '                    Return Me.countdownBrush
    '                End If
    '            End If
    '
    '            ' Llama al procedimiento original
    '            Return NativeMethods.CallWindowProc(Me.previousDlgWndProc, hWnd, msg, wParam, lParam)
    '        End Function
    '
    '        ''' <summary>
    '        ''' Creates and positions the countdown label window inside the dialog window.
    '        ''' <para></para>
    '        ''' It uses native Win32 API calls to create and size the label dynamically based on the expected text size.
    '        ''' </summary>
    '        <DebuggerStepThrough>
    '        Protected Overridable Sub CreateCountdownLabelWindow(dlgHwnd As IntPtr)
    '
    '            Dim dlgClientRect As NativeRectangle
    '            NativeMethods.GetClientRect(dlgHwnd, dlgClientRect)
    '            Dim dlgWidth As Integer = dlgClientRect.Right - dlgClientRect.Left
    '            Const minWidth As Integer = 340
    '
    '            Dim labelWidth As Integer = 100
    '            Dim labelHeight As Integer = 20
    '            Dim labelX As Integer = 0
    '            Dim labelY As Integer = If(dlgWidth < minWidth, 0, dlgClientRect.Bottom - labelHeight)
    '
    '            Dim formatted As String = $"{Math.Floor(Me.TimeOut.TotalHours):00}:{Me.TimeOut.Minutes:00}:{Me.TimeOut.Seconds:00}"
    '            Me.countdownLabelHandle = NativeMethods.CreateWindowEx(
    '                Me.CountdownWindowStyles,
    '                "STATIC",
    '                $" {formatted} ",
    '                WindowStyles.Child Or WindowStyles.Visible,
    '                labelX, labelY,
    '                labelWidth, labelHeight,
    '                dlgHwnd, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero
    '            )
    '
    '            ' Measures the proper label's widow width and height.
    '            Const textToDraw As String = " 99:59:59 "
    '            Dim hdc As IntPtr = NativeMethods.GetDC(Me.countdownLabelHandle)
    '            Dim textSize As NativeSize
    '            NativeMethods.GetTextExtentPoint32(hdc, textToDraw, textToDraw.Length, textSize)
    '            NativeMethods.ReleaseDC(Me.countdownLabelHandle, hdc)
    '            labelWidth = textSize.Width
    '            labelHeight = textSize.Height
    '            labelY = If(dlgWidth < minWidth, 0, dlgClientRect.Bottom - labelHeight)
    '
    '            NativeMethods.SetWindowPos(
    '                Me.countdownLabelHandle,
    '                IntPtr.Zero, labelX, labelY,
    '                labelWidth, labelHeight,
    '                SetWindowPosFlags.IgnoreZOrder Or SetWindowPosFlags.DontActivate Or SetWindowPosFlags.ShowWindow
    '            )
    '
    '        End Sub
    '
    '        ''' <summary>
    '        ''' Cleans-up managed and native resources used by this class.
    '        ''' <para></para>
    '        ''' This method must be called always before calling <see cref="DevMessageBox.ReturnMessageBoxResult"/> function.
    '        ''' </summary>
    '        <DebuggerStepThrough>
    '        Protected Overridable Sub CleanUp()
    '
    '            If Me.CountdownTimer IsNot Nothing Then
    '                Me.CountdownTimer.Stop()
    '                Me.CountdownTimer.Dispose()
    '                Me.CountdownTimer = Nothing
    '            End If
    '
    '            If Me.DestroyTimer IsNot Nothing Then
    '                Me.DestroyTimer.Enabled = False
    '                Me.DestroyTimer.Dispose()
    '                Me.DestroyTimer = Nothing
    '            End If
    '
    '            If Me.transparentForm IsNot Nothing Then
    '                Me.transparentForm.Dispose()
    '                Me.transparentForm = Nothing
    '            End If
    '
    '            If Me.countdownLabelHandle <> IntPtr.Zero Then
    '                NativeMethods.DestroyWindow(Me.countdownLabelHandle)
    '                Me.countdownLabelHandle = IntPtr.Zero
    '            End If
    '
    '            If Me.countdownBrush <> IntPtr.Zero Then
    '                NativeMethods.DeleteObject(Me.countdownBrush)
    '                Me.countdownBrush = IntPtr.Zero
    '            End If
    '
    '            If Me.DialogWindowHandle <> IntPtr.Zero Then
    '                NativeMethods.EndDialog(Me.DialogWindowHandle, IntPtr.Zero) ' DialogResult.None
    '                Me.DialogWindowHandle = IntPtr.Zero
    '            End If
    '
    '        End Sub
    '
    '#End Region
    '
    '#Region " Event Handlers "
    '
    '        ''' <summary>
    '        ''' Handles the <see cref="System.Windows.Forms.Timer.Tick"/> event of the <see cref="DevMessageBox.CountdownTimer"/> control.
    '        ''' </summary>
    '        '''
    '        ''' <param name="sender">
    '        ''' The source of the event.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="e">
    '        ''' The <see cref="EventArgs"/> instance containing the event data.
    '        ''' </param>
    '        <DebuggerStepThrough>
    '        Protected Overridable Sub CountdownTimer_Tick(sender As Object, e As EventArgs) Handles CountdownTimer.Tick
    '
    '            If Me.countdownLabelHandle <> IntPtr.Zero AndAlso Me.countdownSeconds <> 0 Then
    '                Dim ts As TimeSpan = TimeSpan.FromSeconds(Interlocked.Decrement(Me.countdownSeconds))
    '                Dim formatted As String = $"{Math.Floor(ts.TotalHours):00}:{ts.Minutes:00}:{ts.Seconds:00}"
    '                NativeMethods.SetWindowText(Me.countdownLabelHandle, $" {formatted} ")
    '            End If
    '        End Sub
    '
    '        ''' <summary>
    '        ''' Handles the <see cref="System.Windows.Forms.Timer.Tick"/> event of the <see cref="DevMessageBox.DestroyTimer"/> control.
    '        ''' </summary>
    '        '''
    '        ''' <param name="sender">
    '        ''' The source of the event.
    '        ''' </param>
    '        ''' 
    '        ''' <param name="e">
    '        ''' The <see cref="EventArgs"/> instance containing the event data.
    '        ''' </param>
    '        <DebuggerStepThrough>
    '        Protected Overridable Sub DestroyTimer_Tick(sender As Object, e As EventArgs) Handles DestroyTimer.Tick
    '
    '            Me.CleanUp()
    '        End Sub
    '
    '#End Region
    '
    '#Region " IDisposable Implementation "
    '
    '        ''' <summary>
    '        ''' Flag to detect redundant calls when disposing.
    '        ''' </summary>
    '        Private isDisposed As Boolean = False
    '
    '        ''' <summary>
    '        ''' Releases all the resources used by this instance.
    '        ''' </summary>
    '        <DebuggerStepThrough>
    '        Public Sub Dispose() Implements IDisposable.Dispose
    '            Me.Dispose(isDisposing:=True)
    '            GC.SuppressFinalize(Me)
    '        End Sub
    '
    '        ''' <summary>
    '        ''' Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
    '        ''' Releases unmanaged and, optionally, managed resources.
    '        ''' </summary>
    '        '''
    '        ''' <param name="isDisposing">
    '        ''' <see langword="True"/>  to release both managed and unmanaged resources; 
    '        ''' <see langword="False"/> to release only unmanaged resources.
    '        ''' </param>
    '        <DebuggerStepThrough>
    '        Protected Overridable Sub Dispose(isDisposing As Boolean)
    '
    '            If (Not Me.isDisposed) AndAlso isDisposing Then
    '
    '                Me.CleanUp()
    '                Me.enumWindowCurrentTries = 0
    '                Me.countdownSeconds = 0
    '                Me.consoleRect = Rectangle.Empty
    '            End If
    '
    '            Me.isDisposed = True
    '        End Sub
    '
    '#End Region
    '
    '    End Class

End Namespace

#End Region
