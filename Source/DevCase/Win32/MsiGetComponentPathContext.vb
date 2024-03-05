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

#Region " MsiGetComponentPathContext "

Namespace DevCase.Interop.Unmanaged.Win32.Enums

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Flags for <see cref="MsiGetComponentPathExW"/> function.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' <see href="https://learn.microsoft.com/en-us/windows/win32/api/msi/nf-msi-msigetcomponentpathexw"/>
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    Public Enum MsiGetComponentPathContext As Integer

        ''' <summary>
        ''' Include applications installed in the per–user–managed installation context.
        ''' </summary>
        UserManaged = 1

        ''' <summary>
        '''  Include applications installed in the per–user–unmanaged installation context. 
        ''' </summary>
        UserUnmanaged = 2

        ''' <summary>
        ''' Include applications installed in the per-machine installation context. 
        ''' <para></para>
        ''' Note: When parameter 'context' parameter is set to <see cref="MsiGetComponentPathContext.Machine"/> only, 
        ''' the 'userSid' parameter must be NULL.
        ''' </summary>
        Machine = 3

    End Enum

End Namespace

#End Region
