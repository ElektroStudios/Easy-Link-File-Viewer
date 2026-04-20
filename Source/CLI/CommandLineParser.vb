#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Windows.Forms

Imports DevCase.Core.IO

#End Region

#Region " CommandLineParser "

''' 
''' <summary>
''' Provides methods to parse Microsoft-style command-line arguments (e.g. <c>/Switch:Value</c>)
''' into a <see cref="CommandLineOptions"/> instance.
''' </summary>
''' 
Friend NotInheritable Class CommandLineParser

#Region " Constructors "

    ''' 
    ''' <summary>
    ''' Prevents a default instance of the <see cref="CommandLineParser"/> class from being created.
    ''' </summary>
    ''' 
    Private Sub New()
    End Sub

#End Region

#Region " Public Methods "

    ''' 
    ''' <summary>
    ''' Determines whether the specified command-line arguments indicate CLI mode
    ''' (as opposed to simply passing a file path for GUI mode).
    ''' </summary>
    ''' 
    ''' <param name="args">
    ''' The command-line arguments to evaluate.
    ''' </param>
    ''' 
    ''' <returns>
    ''' <see langword="True"/> if the arguments contain at least one switch starting with "/";
    ''' otherwise, <see langword="False"/>.
    ''' </returns>
    ''' 
    Friend Shared Function IsCliMode(args As IEnumerable(Of String)) As Boolean

        For Each arg As String In args
            If arg.StartsWith("/", StringComparison.Ordinal) Then
                Return True
            End If
        Next
        Return False
    End Function

    ''' 
    ''' <summary>
    ''' Parses the specified command-line arguments into a <see cref="CommandLineOptions"/> instance.
    ''' </summary>
    ''' 
    ''' <param name="args">
    ''' The command-line arguments to parse.
    ''' </param>
    ''' 
    ''' <returns>
    ''' A <see cref="CommandLineOptions"/> instance populated with the parsed values.
    ''' </returns>
    ''' 
    ''' <exception cref="ArgumentException">
    ''' Thrown when an unknown switch or invalid value is encountered.
    ''' </exception>
    ''' 
    Friend Shared Function Parse(args As IList(Of String)) As CommandLineOptions

        Dim options As New CommandLineOptions()

        For i As Integer = 0 To args.Count - 1
            Dim arg As String = args(i)

            If Not arg.StartsWith("/", StringComparison.Ordinal) Then
                ' Treat as a positional file path if /File was not explicitly specified.
                If String.IsNullOrEmpty(options.FilePath) Then
                    options.FilePath = arg
                End If
                Continue For
            End If

            Dim switchName As String
            Dim switchValue As String = Nothing

            Dim colonIndex As Integer = arg.IndexOf(":"c)
            If colonIndex > 0 Then
                switchName = arg.Substring(1, colonIndex - 1)
                switchValue = arg.Substring(colonIndex + 1)
            Else
                switchName = arg.Substring(1)
            End If

            Select Case switchName.ToUpperInvariant()

                ' Action switches.
                Case "CREATE"
                    CommandLineParser.SetAction(options, CommandLineAction.Create)

                Case "MODIFY"
                    CommandLineParser.SetAction(options, CommandLineAction.Modify)

                Case "READ"
                    CommandLineParser.SetAction(options, CommandLineAction.Read)

                Case "HELP", "?"
                    CommandLineParser.SetAction(options, CommandLineAction.Help)

                ' Flag switches.
                Case "FORCE"
                    options.Force = True

                ' Property switches that require a value.
                Case "FILE"
                    CommandLineParser.RequireValue(switchName, switchValue)
                    options.FilePath = switchValue

                Case "TARGET"
                    CommandLineParser.RequireValue(switchName, switchValue)
                    options.Target = switchValue

                Case "ARGS"
                    CommandLineParser.RequireValue(switchName, switchValue)
                    options.TargetArguments = switchValue

                Case "DESC"
                    CommandLineParser.RequireValue(switchName, switchValue)
                    options.Description = switchValue

                Case "ICON"
                    CommandLineParser.RequireValue(switchName, switchValue)
                    options.Icon = switchValue

                Case "ICONINDEX"
                    CommandLineParser.RequireValue(switchName, switchValue)
                    Dim index As Integer = 0
                    If Not Integer.TryParse(switchValue, NumberStyles.Integer, CultureInfo.InvariantCulture, index) Then
                        Throw New ArgumentException($"Invalid /IconIndex value: {switchValue}. Expected an integer.")
                    End If
                    options.IconIndex = index

                Case "HOTKEY"
                    CommandLineParser.RequireValue(switchName, switchValue)
                    options.Hotkey = CommandLineParser.ParseHotkey(switchValue)

                Case "WINDOWSTATE"
                    CommandLineParser.RequireValue(switchName, switchValue)
                    options.WindowState = CommandLineParser.ParseWindowState(switchValue)

                Case "WORKDIR"
                    CommandLineParser.RequireValue(switchName, switchValue)
                    options.WorkingDirectory = switchValue

                Case "APPID"
                    CommandLineParser.RequireValue(switchName, switchValue)
                    options.AppId = switchValue

                Case Else
                    Throw New ArgumentException($"Unknown switch: /{switchName}. Use /Help for usage information.")

            End Select
        Next

        Return options

    End Function

#End Region

#Region " Private Methods "

    ''' 
    ''' <summary>
    ''' Sets the action on the specified options, throwing if an action was already specified.
    ''' </summary>
    ''' 
    ''' <param name="options">
    ''' The options instance to update.
    ''' </param>
    ''' 
    ''' <param name="action">
    ''' The action to set.
    ''' </param>
    ''' 
    ''' <exception cref="ArgumentException">
    ''' Thrown when an action has already been specified.
    ''' </exception>
    ''' 
    Private Shared Sub SetAction(options As CommandLineOptions, action As CommandLineAction)

        If options.Action <> CommandLineAction.None Then
            Throw New ArgumentException(
                $"Multiple actions specified: /{options.Action} and /{action}. Only one action is allowed.")
        End If
        options.Action = action
    End Sub

    ''' 
    ''' <summary>
    ''' Validates that a switch value was provided (i.e. the switch was written as <c>/Switch:value</c>).
    ''' </summary>
    ''' 
    ''' <param name="switchName">
    ''' The name of the switch, for the error message.
    ''' </param>
    ''' 
    ''' <param name="switchValue">
    ''' The value to validate.
    ''' </param>
    ''' 
    ''' <exception cref="ArgumentException">
    ''' Thrown when <paramref name="switchValue"/> is <see langword="Nothing"/>,
    ''' meaning the switch was written without a colon and value.
    ''' </exception>
    ''' 
    Private Shared Sub RequireValue(switchName As String, switchValue As String)

        If switchValue Is Nothing Then
            Throw New ArgumentException($"/{switchName} requires a value. Use /{switchName}:value syntax.")
        End If
    End Sub

    ''' 
    ''' <summary>
    ''' Parses a hotkey string value into a <see cref="Keys"/> value.
    ''' <para></para>
    ''' Supports numeric values (e.g. <c>131184</c>), single key names (e.g. <c>F1</c>),
    ''' and combined keys using "+" or "," as separator (e.g. <c>Control+Alt+F1</c>).
    ''' </summary>
    ''' 
    ''' <param name="value">
    ''' The hotkey string to parse.
    ''' </param>
    ''' 
    ''' <returns>
    ''' The parsed <see cref="Keys"/> value.
    ''' </returns>
    ''' 
    ''' <exception cref="ArgumentException">
    ''' Thrown when the value cannot be parsed as a valid hotkey.
    ''' </exception>
    ''' 
    Private Shared Function ParseHotkey(value As String) As Keys

        ' Try numeric value first (e.g. "131184").
        Dim numericValue As Integer = 0
        If Integer.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, numericValue) Then
            Return CType(numericValue, Keys)
        End If

        ' Normalize "+" separators to "," for Enum.Parse compatibility
        ' (e.g. "Control+Alt+F1" becomes "Control, Alt, F1").
        Dim normalizedValue As String = value.Replace("+", ", ")

        Dim result As Keys = Keys.None
        If [Enum].TryParse(Of Keys)(normalizedValue, True, result) Then
            Return result
        End If

        Throw New ArgumentException(
            $"Invalid /Hotkey value: {value}. " &
            $"Use key names separated by '+' (e.g. Control+Alt+F1), 'None', or a numeric value.")

    End Function

    ''' 
    ''' <summary>
    ''' Parses a window state string value into a <see cref="ShortcutWindowState"/> value.
    ''' <para></para>
    ''' Supports enum names (<c>Normal</c>, <c>Minimized</c>, <c>Maximized</c>) and
    ''' common aliases (<c>Min</c>, <c>Max</c>).
    ''' </summary>
    ''' 
    ''' <param name="value">
    ''' The window state string to parse.
    ''' </param>
    ''' 
    ''' <returns>
    ''' The parsed <see cref="ShortcutWindowState"/> value.
    ''' </returns>
    ''' 
    ''' <exception cref="ArgumentException">
    ''' Thrown when the value cannot be parsed as a valid window state.
    ''' </exception>
    ''' 
    Private Shared Function ParseWindowState(value As String) As ShortcutWindowState

        Select Case value.ToUpperInvariant()
            Case "MIN"
                Return ShortcutWindowState.Minimized
            Case "MAX"
                Return ShortcutWindowState.Maximized
        End Select

        Dim result As ShortcutWindowState = ShortcutWindowState.Normal
        If [Enum].TryParse(Of ShortcutWindowState)(value, True, result) Then
            Return result
        End If

        Throw New ArgumentException(
            $"Invalid /WindowState value: {value}. Valid values: Normal, Minimized (Min), Maximized (Max).")

    End Function

#End Region

End Class

#End Region