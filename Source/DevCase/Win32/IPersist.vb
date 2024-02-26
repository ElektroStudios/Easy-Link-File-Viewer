' This source-code is freely distributed as part of "DevCase for .NET Framework".
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

#Region " Imports "

Imports System.Runtime.InteropServices

#End Region

#Region " IPersist "

Namespace DevCase.Interop.Unmanaged.Win32.Interfaces

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Provides the CLSID of an object that can be stored persistently in the system.
    ''' <para></para>
    ''' Allows the object to specify which object handler to use in the client process, 
    ''' as it is used in the default implementation of marshaling.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms688695%28v=vs.85%29.aspx"/>
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    <ComImport>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <Guid("0000010c-0000-0000-c000-000000000046")>
    Friend Interface IPersist

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves the class identifier (CLSID) of the object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="refClassID">
        ''' A pointer to the location that receives the CLSID on return.
        ''' <para></para>
        ''' The CLSID is a globally unique identifier (GUID) that uniquely represents an object class that 
        ''' defines the code that can manipulate the object's data.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <PreserveSig()>
        Sub GetClassID(ByRef refClassID As Guid)

    End Interface

End Namespace

#End Region
