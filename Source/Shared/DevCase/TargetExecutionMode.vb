Imports System.ComponentModel

Namespace DevCase.Core.IO

    ''' <summary>
    ''' Defines the target execution strategy and path resolution mode for a shortcut file.
    ''' </summary>
    Public Enum ShortcutTargetExecutionMode As Integer

        ''' <summary>
        ''' The default Windows behavior. The shortcut target is stored and executed "as is".
        ''' </summary>
        <Description("Standard Mode")>
        Standard = 0

        ''' <summary>
        ''' Calculates a relative path to the target and routes the execution through the Windows Explorer shell.
        ''' <para></para>
        ''' Ideal for simple portable targets that do not require additional command-line arguments.
        ''' </summary>
        <Description("Relative Path via Windows Explorer")>
        RelativePathViaExplorer = 1

        ''' <summary>
        ''' Calculates a relative path to the target and routes the execution through the Command Prompt.
        ''' <para></para>
        ''' Necessary for portable targets that must also receive their own command-line arguments.
        ''' </summary>
        <Description("Relative Path with Arguments via Command Prompt")>
        RelativePathWithArgumentsViaCommandPrompt = 2

    End Enum

End Namespace
