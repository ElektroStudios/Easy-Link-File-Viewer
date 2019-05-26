' ***********************************************************************
' Author   : ElektroStudios
' Modified : 10-November-2015
' ***********************************************************************

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " IShellLink.Resolve Flags "

Namespace DevCase.Interop.Unmanaged.Win32.Enums

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' IShellLink.Resolve method action flags. 
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/bb774952%28v=vs.85%29.aspx"/>
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    <Flags>
    Public Enum IShellLinkResolveFlags

        ''' <summary>
        ''' Do not display a dialog box if the link cannot be resolved. When SLR_NO_UI is set,
        ''' the high-order word of fFlags can be set to a time-out value that specifies the
        ''' maximum amount of time to be spent resolving the link. The function returns if the
        ''' link cannot be resolved within the time-out duration. If the high-order word is set
        ''' to zero, the time-out duration will be set to the default value of 3,000 milliseconds
        ''' (3 seconds). To specify a value, set the high word of fFlags to the desired time-out
        ''' duration, in milliseconds.
        ''' </summary>
        NoUI = &H1

        ''' <summary>
        ''' If the link object has changed, update its path and list of identifiers.
        ''' If SLR_UPDATE is set, you do not need to call IPersistFile::IsDirty to determine,
        ''' whether or not the link object has changed.
        ''' </summary>
        Update = &H4

        ''' <summary>
        ''' Do not update the link information
        ''' </summary>
        NoUpdate = &H8

        ''' <summary>
        ''' Do not execute the search heuristics
        ''' </summary>
        NoSearch = &H10

        ''' <summary>
        ''' Do not use distributed link tracking
        ''' </summary>
        Notrack = &H20

        ''' <summary>
        ''' Disable distributed link tracking. 
        ''' By default, distributed link tracking tracks removable media,
        ''' across multiple devices based on the volume name. 
        ''' It also uses the Universal Naming Convention (UNC) path to track remote file systems,
        ''' whose drive letter has changed.
        ''' Setting SLR_NOLINKINFO disables both types of tracking.
        ''' </summary>
        NoLinkInfo = &H40

        ''' <summary>
        ''' Call the Microsoft Windows Installer
        ''' </summary>
        InvokeMsi = &H80

    End Enum

End Namespace

#End Region
