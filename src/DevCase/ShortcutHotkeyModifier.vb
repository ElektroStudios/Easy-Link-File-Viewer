' ***********************************************************************
' Author   : ElektroStudios
' Modified : 14-November-2015
' ***********************************************************************

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Hotkey Modifier "

Namespace DevCase.Core.IO

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Specifies a key-modifier to assign for a shortcut.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/bb774926%28v=vs.85%29.aspx"/>
    ''' <para></para>
    ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms646278%28v=vs.85%29.aspx"/>
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    <Flags>
    Public Enum ShortcutHotkeyModifier As Short

        ''' <summary>
        ''' Specifies any modifier.
        ''' </summary>
        None = 0

        ''' <summary>
        ''' The <c>SHIFT</c> keyboard key.
        ''' </summary>
        Shift = 1

        ''' <summary>
        ''' The <c>CTRL</c> keyboard key.
        ''' </summary>
        Control = 2

        ''' <summary>
        ''' The <c>ALT</c> keyboard key.
        ''' </summary>
        Alt = 4

    End Enum

End Namespace

#End Region
