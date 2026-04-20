' ***********************************************************************
' Author   : ElektroStudios
' Modified : 08-October-2024
' ***********************************************************************

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.Runtime.InteropServices
Imports System.Security

#End Region

#Region " P/Invoking "

' ReSharper disable once CheckNamespace

Namespace DevCase.Win32.Delegates

    ''' <summary>
    ''' Contains delegate callbacks for Win32 functions.
    ''' </summary>
    <HideModuleName>
    Friend Module Delegates

#Region " Delegates "

        ''' <summary>
        ''' An application-defined callback function used with the <see cref="NativeMethods.EnumThreadWindows"/> function.
        ''' <para></para>
        ''' It receives the window handles associated with a thread.
        ''' <para></para>
        ''' The <c>WNDENUMPROC</c> type defines a pointer to this callback function.
        ''' <para></para>
        ''' <see cref="EnumThreadWindowsProc"/> is a placeholder for the application-defined function name. 
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms633496%28v=vs.85%29.aspx"/>
        ''' </remarks>
        '''
        ''' <param name="hWnd">
        ''' A handle to a window associated with the thread specified in the <see cref="NativeMethods.EnumThreadWindows"/> function. 
        ''' </param>
        ''' 
        ''' <param name="lParam">
        ''' The application-defined value given in <see cref="NativeMethods.EnumThreadWindows"/> functions.
        ''' </param>
        '''
        ''' <returns>
        ''' To continue enumeration, the callback function must return <see langword="True"/>; 
        ''' to stop enumeration, it must return <see langword="False"/>. 
        ''' </returns>
        <SuppressUnmanagedCodeSecurity>
        Friend Delegate Function EnumThreadWindowsProc(hWnd As IntPtr,
                                                       lParam As IntPtr
       ) As Boolean

        ''' <summary>
        ''' A callback function, which you define in your application, that processes messages sent to a window. 
        ''' <para></para>
        ''' The WNDPROC type defines a pointer to this callback function. 
        ''' The WndProc name is a placeholder for the name of the function that you define in your application.
        ''' </summary>
        '''
        ''' <remarks>
        ''' <see href="https://learn.microsoft.com/en-us/windows/win32/api/winuser/nc-winuser-wndproc"/>
        ''' </remarks>
        '''
        ''' <param name="hWnd">
        ''' A handle to the window. This parameter is typically named <c>hWnd</c>.
        ''' </param>
        '''
        ''' <param name="msg">
        ''' The message. This parameter is typically named <c>msg</c>.
        ''' <para></para>
        ''' For lists of the system-provided messages, see <see cref="WindowMessages"/>.
        ''' </param>
        '''
        ''' <param name="wParam">
        ''' Additional message information. This parameter is typically named <c>wParam</c>.
        ''' <para></para>
        ''' The contents of the wParam parameter depend on the value of the <paramref name="msg"/> parameter.
        ''' </param>
        '''
        ''' <param name="lParam">
        ''' Additional message information. This parameter is typically named <c>lParam</c>.
        ''' <para></para>
        ''' The contents of the lParam parameter depend on the value of the <paramref name="msg"/> parameter.
        ''' </param>
        '''
        ''' <returns>
        ''' The return value is the result of the message processing, and depends on the message sent.
        ''' </returns>
        <UnmanagedFunctionPointer(CallingConvention.Winapi, CharSet:=CharSet.Auto, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Delegate Function WndProc(hWnd As IntPtr,
                                         msg As UInteger,
                                         wParam As IntPtr,
                                         lParam As IntPtr
                                        ) As IntPtr

#End Region

    End Module

End Namespace

#End Region
