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

#Region " IPersistFile "

Namespace DevCase.Interop.Unmanaged.Win32.Interfaces

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Enables an object to be loaded from or saved to a disk file, rather than a storage object or stream.
    ''' <para></para>
    ''' Because the information needed to open a file varies greatly from one application to another, 
    ''' the implementation of <see cref="IPersistFile.Load"/> on the object must also open its disk file.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms687223%28v=vs.85%29.aspx"/>
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    <ComImport>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <Guid("0000010b-0000-0000-C000-000000000046")>
    Friend Interface IPersistFile : Inherits IPersist

        Shadows Sub GetClassID(ByRef refClassID As Guid)

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Determines whether an object has changed since it was last saved to its current file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' This method returns <see cref="Enums.HResult.S_OK"/> to indicate that the object has changed. Otherwise, it returns <c>S_FALSE</c>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <PreserveSig()>
        Function IsDirty() As Integer

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Opens the specified file and initializes an object from the file contents.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="fileName">
        ''' The absolute path of the file to be opened.
        ''' </param>
        ''' 
        ''' <param name="mode">
        ''' The access mode to be used when opening the file.
        ''' <para></para>
        ''' Possible values are taken from the <c>STGM</c> enumeration.
        ''' <para></para>
        ''' The method can treat this value as a suggestion, adding more restrictive permissions if necessary.
        ''' <para></para>
        ''' If <paramref name="mode"/> is <c>0</c>, 
        ''' the implementation should open the file using whatever default permissions are used when a user opens the file.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <PreserveSig()>
        Sub Load(<[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal fileName As String,
                                                           ByVal mode As UInteger)

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Saves a copy of the object to the specified file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="fileName">
        ''' The absolute path of the file to which the object should be saved.
        ''' <para></para>
        ''' If <paramref name="fileName"/> is <see langword="Nothing"/>, 
        ''' the object should save its data to the current file, if there is one.
        ''' </param>
        ''' 
        ''' <param name="remember">
        ''' Indicates whether the <paramref name="fileName"/> parameter is to be used as the current working file.
        ''' <para></para>
        ''' If <see langword="True"/>, <paramref name="fileName"/> becomes the current file and the 
        ''' object should clear its dirty flag after the save.
        ''' <para></para>
        ''' If <see langword="False"/>, this save operation is a <c>Save A Copy As</c> ... operation. 
        ''' In this case, the current file is unchanged and the object should not clear its dirty flag.
        ''' <para></para>
        ''' If <paramref name="fileName"/> is <see langword="Nothing"/>, 
        ''' the implementation should ignore the <paramref name="fRemember"/> flag.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <PreserveSig()>
        Sub Save(<[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal fileName As String,
                 <[In](), MarshalAs(UnmanagedType.Bool)> ByVal remember As Boolean)

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the object that it can write to its file.
        ''' <para></para>
        ''' It does this by notifying the object that it can revert from <c>NoScribble</c> mode 
        ''' (in which it must not write to its file), to <c>Normal</c> mode (in which it can).
        ''' <para></para>
        ''' The component enters <c>NoScribble</c> mode when it receives an <see cref="IPersistFile.Save"/> call.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="fileName">
        ''' The absolute path of the file where the object was saved previously.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <PreserveSig()>
        Sub SaveCompleted(<[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal fileName As String)

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves the current name of the file associated with the object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="fileName">
        ''' The path for the current file or the default file name prompt (such as <c>*.txt</c>).
        ''' <para></para>
        ''' If an error occurs, <paramref name="fileName"/> is set to <see langword="Nothing"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <PreserveSig()>
        Sub GetCurFile(<[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal fileName As String)

    End Interface

End Namespace

#End Region
