#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.Windows.Forms

Imports DevCase.Core.IO

#End Region

#Region " CommandLineOptions "

''' 
''' <summary>
''' Represents the parsed command-line options for the application.
''' <para></para>
''' Properties set to <see langword="Nothing"/> indicate they were not specified on the command line.
''' This distinction is important for the <see cref="CommandLineAction.Modify"/> action, where only
''' explicitly specified properties should be changed.
''' </summary>
''' 
Friend NotInheritable Class CommandLineOptions

#Region " Properties "

    ''' 
    ''' <summary>
    ''' Gets or sets the action to perform.
    ''' </summary>
    ''' 
    Friend Property Action As CommandLineAction = CommandLineAction.None

    ''' 
    ''' <summary>
    ''' Gets or sets the file path of the shortcut (.lnk) file.
    ''' </summary>
    ''' 
    Friend Property FilePath As String = String.Empty

    ''' 
    ''' <summary>
    ''' Gets or sets the target path for the shortcut.
    ''' </summary>
    ''' 
    Friend Property Target As String = Nothing

    ''' 
    ''' <summary>
    ''' Gets or sets the target command-line arguments for the shortcut.
    ''' </summary>
    ''' 
    Friend Property TargetArguments As String = Nothing

    ''' 
    ''' <summary>
    ''' Gets or sets the description of the shortcut.
    ''' </summary>
    ''' 
    Friend Property Description As String = Nothing

    ''' 
    ''' <summary>
    ''' Gets or sets the icon file path for the shortcut.
    ''' </summary>
    ''' 
    Friend Property Icon As String = Nothing

    ''' 
    ''' <summary>
    ''' Gets or sets the icon index within the icon file.
    ''' </summary>
    ''' 
    Friend Property IconIndex As Nullable(Of Integer) = Nothing

    ''' 
    ''' <summary>
    ''' Gets or sets the hotkey assigned to the shortcut.
    ''' </summary>
    ''' 
    Friend Property Hotkey As Nullable(Of Keys) = Nothing

    ''' 
    ''' <summary>
    ''' Gets or sets the window state for the shortcut target.
    ''' </summary>
    ''' 
    Friend Property WindowState As Nullable(Of ShortcutWindowState) = Nothing

    ''' 
    ''' <summary>
    ''' Gets or sets the working directory for the shortcut.
    ''' </summary>
    ''' 
    Friend Property WorkingDirectory As String = Nothing

    ''' 
    ''' <summary>
    ''' Gets or sets the application user model ID for the shortcut.
    ''' </summary>
    ''' 
    Friend Property AppId As String = Nothing

    ''' 
    ''' <summary>
    ''' Gets or sets a value indicating whether to overwrite an existing file during creation.
    ''' </summary>
    ''' 
    Friend Property Force As Boolean = False

#End Region

End Class

#End Region