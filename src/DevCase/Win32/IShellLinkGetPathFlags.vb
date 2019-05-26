' ***********************************************************************
' Author   : ElektroStudios
' Modified : 10-November-2015
' ***********************************************************************

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " IShellLink.GetPath Flags "

Namespace DevCase.Interop.Unmanaged.Win32.Enums

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' IShellLink.GetPath method flags that specify the type of path information to retrieve. 
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/bb774944%28v=vs.85%29.aspx"/>
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    <Flags>
    Public Enum IShellLinkGetPathFlags

        ''' <summary>
        ''' Retrieves the standard short (8.3 format) file name.
        ''' </summary>
        ShortPath = &H1

        ''' <summary>
        ''' Retrieves the Universal Naming Convention (UNC) path name of the file.
        ''' </summary>
        UncPriority = &H2

        ''' <summary>
        ''' Retrieves the raw path name. 
        ''' A raw path is something that might not exist and may include environment variables that need to be expanded.
        ''' </summary>
        RawPath = &H4

    End Enum

End Namespace

#End Region
