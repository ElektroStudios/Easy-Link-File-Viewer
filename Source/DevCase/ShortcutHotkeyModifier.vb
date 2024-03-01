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
    Friend Enum ShortcutHotkeyModifier As Short

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
