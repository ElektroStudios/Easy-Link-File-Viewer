#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " CommandLineAction "

''' 
''' <summary>
''' Specifies the action to perform when the application runs in command-line mode.
''' </summary>
''' 
Friend Enum CommandLineAction As Integer

    ''' <summary>
    ''' No action specified. The application should run in GUI mode.
    ''' </summary>
    None

    ''' <summary>
    ''' Creates a new shortcut (.lnk) file.
    ''' </summary>
    Create

    ''' <summary>
    ''' Modifies an existing shortcut (.lnk) file.
    ''' </summary>
    Modify

    ''' <summary>
    ''' Reads and displays the properties of a shortcut (.lnk) file.
    ''' </summary>
    Read

    ''' <summary>
    ''' Displays the command-line usage help text.
    ''' </summary>
    Help

End Enum

#End Region
