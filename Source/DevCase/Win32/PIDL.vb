' Unused class. It may be helpful in the future.
#If debug

' ***********************************************************************
' Author   : ElektroStudios
' Modified : 01-June-2019
' ***********************************************************************

#Region " Option Statements "

Option Strict Off
Option Explicit On
Option Infer Off

#End Region

#Region " Usage Examples "

#Region " Example #1 "

' Dim pidl As New PIDL("C:\Windows")
' Console.WriteLine(pidl.ToString(ShellItemGetDisplayName.DesktopAbsoluteParsing))
' pidl.Dispose()

#End Region

#Region " Example #2 "

' Dim pidl As PIDL = NativeMethods.ILCreateFromPath("C:\Windows")
' Console.WriteLine(pidl.ToString(ShellItemGetDisplayName.DesktopAbsoluteParsing))
' pidl.Dispose()

#End Region

#End Region

#Region " Imports "

Imports System.Runtime.ConstrainedExecution
Imports System.Runtime.InteropServices
Imports System.Text
Imports Microsoft.Win32.SafeHandles

Imports DevCase.Interop.Unmanaged.Win32.Enums
Imports DevCase.Interop.Unmanaged.Win32

#End Region

#Region " PIDL "

' ReSharper disable once CheckNamespace

Namespace DevCase.Win32.Common

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Represents a managed pointer to an ITEMIDLIST structure.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' ITEMIDLIST documentation:
    ''' <see href="https://docs.microsoft.com/en-us/windows/desktop/api/shtypes/ns-shtypes-_itemidlist"/>
    ''' <para></para>
    ''' Original source-code:
    ''' <see href="https://github.com/dahall/Vanara/blob/master/PInvoke/Shell32/ShTypes.PIDL.cs"/>
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code language="VB.NET">
    ''' Dim pidl As New PIDL("C:\Windows")
    ''' Console.WriteLine(pidl.ToString(ShellItemGetDisplayName.DesktopAbsoluteParsing))
    ''' pidl.Dispose()
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code language="VB.NET">
    ''' Dim pidl As PIDL = NativeMethods.ILCreateFromPath("C:\Windows")
    ''' Console.WriteLine(pidl.ToString(ShellItemGetDisplayName.DesktopAbsoluteParsing))
    ''' pidl.Dispose()
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    <CodeAnalysis.SuppressMessage("CodeQuality", "IDE0079:Remove unnecessary suppression", Justification:="Required to migrate this code to .NET Core")>
    Public NotInheritable Class PIDL : Inherits SafeHandleZeroOrMinusOneIsInvalid : Implements IEnumerable(Of PIDL),
                                                                                               IEquatable(Of PIDL),
                                                                                               IEquatable(Of IntPtr)

#Region " Properties "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a value indicating whether this ITEMIDLIST is empty.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if this ITEMIDLIST is empty; otherwise, <see langword="False"/>.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public ReadOnly Property IsEmpty As Boolean
            Get
                Return PIDL.ILIsEmpty(Me.handle)
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the last SHITEMID in this ITEMIDLIST.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The last SHITEMID.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public ReadOnly Property LastId As PIDL
            Get
                Return New PIDL(NativeMethods.ILFindLastID(Me.handle), True)
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets an ITEMIDLIST with the last ID removed. 
        ''' If this is the topmost ID, a clone of the current is returned.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public ReadOnly Property Parent As PIDL
            Get
                Dim p As New PIDL(Me)
                p.RemoveLastId()
                Return p
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the size, in bytes, of this ITEMIDLIST.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public ReadOnly Property Size As UInteger
            Get
                Return NativeMethods.ILGetSize(Me.handle)
            End Get
        End Property

#End Region

#Region " Constructors "

#Disable Warning CA1419 ' Provide a parameterless constructor that is as visible as the containing type for concrete types derived from 'System.Runtime.InteropServices.SafeHandle'
        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Prevents a default instance of the <see cref="PIDL"/> class from being created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub New()
            MyBase.New(False)
        End Sub
#Enable Warning CA1419 ' Provide a parameterless constructor that is as visible as the containing type for concrete types derived from 'System.Runtime.InteropServices.SafeHandle'

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Initializes a new instance of the <see cref="PIDL"/> class.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidl">
        ''' A handle to a native ITEMIDLIST.
        ''' </param>
        ''' 
        ''' <param name="clone">
        ''' If set to <see langword="True"/>, clone the ITEMIDLIST before storing it.
        ''' </param>
        ''' 
        ''' <param name="ownsHandle">
        ''' If set to <see langword="True"/>, this <see cref="PIDL"/> instance will release the 
        ''' memory associated with the underlying ITEMIDLIST when done.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Public Sub New(pidl As IntPtr, Optional clone As Boolean = False, Optional ownsHandle As Boolean = True)
            MyBase.New(clone OrElse ownsHandle)
            Me.SetHandle(If(clone, NativeMethods.ILCloneIntPtr(pidl), pidl))
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Initializes a new instance of the <see cref="PIDL"/> class.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidl">
        ''' An existing <see cref="PIDL"/> that will be copied and managed.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Public Sub New(pidl As PIDL)
            Me.New(pidl.handle, ownsHandle:=True)
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Initializes a new instance of the <see cref="PIDL"/> class from a file or folder path.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="path">
        ''' A string that specifies the path.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Public Sub New(path As String)
            Me.New(NativeMethods.ILCreateFromPathIntPtr(path))
        End Sub

#End Region

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Appends the specified <see cref="PIDL"/> to the current ITEMIDLIST.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidl">
        ''' The <see cref="PIDL"/> to append.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Public Sub Append(pidl As PIDL)
            Dim newPidl As IntPtr = NativeMethods.ILCombineIntPtr(Me.handle, pidl.DangerousGetHandle())
            Me.ReleaseHandle()
            Me.SetHandle(newPidl)
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Combines the specified <see cref="PIDL"/> instances to create a new one.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="first">
        ''' The first <see cref="PIDL"/>.
        ''' </param>
        ''' 
        ''' <param name="second">
        ''' The second <see cref="PIDL"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A managed <see cref="PIDL"/> instance that contains both supplied lists in their respective order.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Function Combine(first As PIDL, second As PIDL) As PIDL
            Return NativeMethods.ILCombine(first.handle, second.handle)
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Inserts the specified <see cref="PIDL"/> before the current ITEMIDLIST.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidl">
        ''' The <see cref="PIDL"/> to insert.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Public Sub Insert(pidl As PIDL)
            Dim newPidl As IntPtr = NativeMethods.ILCombineIntPtr(pidl.handle, Me.handle)
            Me.ReleaseHandle()
            Me.SetHandle(newPidl)
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Determines if this instance is a parent or ancestor of the specified <see cref="PIDL"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="childPidl">
        ''' Child <see cref="PIDL"/> instance to test.
        ''' </param>
        ''' 
        ''' <param name="immediate">
        ''' If <see langword="True"/>, narrows test to immediate children only.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see langword="True"/> if this instance is a parent or ancestor of the specified <see cref="PIDL"/>.
        ''' <see langword="False"/> otherwise.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Function IsParentOf(childPidl As PIDL, Optional immediate As Boolean = True) As Boolean
            Return NativeMethods.ILIsParent(CType(Me, IntPtr), CType(childPidl, IntPtr), immediate)
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Removes the last identifier from the ITEMIDLIST.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Function RemoveLastId() As Boolean
            Return NativeMethods.ILRemoveLastID(Me.handle)
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Indicates whether the current object is equal to another object of the same type.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="other">
        ''' An object to compare with this object.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see langword="True"/> if the current object is equal to the <paramref name="other"/> parameter; 
        ''' otherwise, <see langword="False"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Overloads Function Equals(other As IntPtr) As Boolean Implements IEquatable(Of IntPtr).Equals
            Return NativeMethods.ILIsEqual(Me.handle, other)
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Indicates whether the current object is equal to another object of the same type.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="other">
        ''' An object to compare with this object.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see langword="True"/> if the current object is equal to the <paramref name="other"/> parameter; 
        ''' otherwise, <see langword="False"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Overloads Function Equals(other As PIDL) As Boolean Implements IEquatable(Of PIDL).Equals
            Return Me.Equals(If(other?.handle, IntPtr.Zero))
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Returns an enumerator that iterates through the collection.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A <see cref="IEnumerator(Of PIDL)"/> that can be used to iterate through the collection.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Function GetEnumerator() As IEnumerator(Of PIDL) Implements IEnumerable(Of PIDL).GetEnumerator
            Return New PIDLEnumerator(Me.handle)
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Returns a <see cref="String"/> that represents this <see cref="PIDL"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A <see cref="String"/> that represents this instance.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Overrides Function ToString() As String
            Return Me.ToString(ShellItemGetDisplayName.NormalDisplay)
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Returns a <see cref="String"/> that represents this <see cref="PIDL"/> 
        ''' according to the format provided by <paramref name="displayNameFormat"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="displayNameFormat">
        ''' The desired display name format.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A <see cref="String"/> that represents this instance.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Overloads Function ToString(displayNameFormat As ShellItemGetDisplayName) As String
            Try
                Dim name As New StringBuilder(260)
                NativeMethods.SHGetNameFromIDList(Me, displayNameFormat, name)
                Return name?.ToString()
            Catch ex As Exception
            End Try
            Return Nothing
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return DirectCast(Me, IEquatable(Of PIDL)).Equals(TryCast(obj, PIDL))
        End Function

        Public Overrides Function GetHashCode() As Integer
            Throw New NotImplementedException
        End Function

#End Region

#Region " Private Methods "

#Disable Warning SYSLIB0004 ' Type or member is obsolete
        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Release the handle
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see langword="True"/> if handle is released successfully, <see langword="False"/> otherwise.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)>
        Protected Overrides Function ReleaseHandle() As Boolean
            NativeMethods.ILFree(Me.handle)
            Return True
        End Function
#Enable Warning SYSLIB0004 ' Type or member is obsolete

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Returns an enumerator that iterates through the collection.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' An <see cref="IEnumerator"/> object that can be used to iterate through the collection.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return Me.GetEnumerator()
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves the next SHITEMID structure in an ITEMIDLIST structure.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidl">
        ''' A constant, unaligned, relative PIDL for which the next SHITEMID structure is being retrieved.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' When this function returns, contains one of three results: 
        ''' If pidl is valid and not the last SHITEMID in the ITEMIDLIST, 
        ''' then it contains a pointer to the next ITEMIDLIST structure.
        ''' If the last ITEMIDLIST structure is passed, it contains NULL, 
        ''' which signals the end of the PIDL. 
        ''' For other values of pidl, the return value is meaningless.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Friend Shared Function ILNext(pidl As IntPtr) As IntPtr
            Dim size As Integer = ILGetItemSize(pidl)
            Return If(size = 0, IntPtr.Zero, New IntPtr(pidl.ToInt64() + size))
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Verifies whether a pointer to an item identifier list (PIDL) is a child PIDL, 
        ''' which is a PIDL with exactly one SHITEMID.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidl">
        ''' A constant, unaligned, relative PIDL that is being checked.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' Returns <see langword="True"/> if the given PIDL is a child PIDL; 
        ''' otherwise, <see langword="False"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Friend Shared Function ILIsChild(pidl As IntPtr) As Boolean
            Return ILIsEmpty(pidl) OrElse ILIsEmpty(ILNext(pidl))
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Returns the size, in bytes, of an SHITEMID structure.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidl">
        ''' A pointer to an SHITEMID structure.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The size of the SHITEMID structure specified by pidl, in bytes.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Friend Shared Function ILGetItemSize(pidl As IntPtr) As Integer
            Return If(pidl.Equals(IntPtr.Zero), 0, Marshal.ReadInt16(pidl))
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Verifies whether an ITEMIDLIST structure is empty.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidl">
        ''' A pointer to the ITEMIDLIST structure to be checked.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see langword="True"/> if the pidl parameter is NULL or the ITEMIDLIST structure pointed to by pidl is empty; 
        ''' otherwise <see langword="False"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Friend Shared Function ILIsEmpty(pidl As IntPtr) As Boolean
            Return ILGetItemSize(pidl) = 0
        End Function

#End Region

#Region " Operators "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Performs an explicit conversion from <see cref="PIDL"/> to <see cref="IntPtr"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pidl">
        ''' The source <see cref="PIDL"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The result of the conversion.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Narrowing Operator CType(pidl As PIDL) As IntPtr
            Return pidl.handle
        End Operator

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Performs an implicit conversion from <see cref="IntPtr"/> to <see cref="PIDL"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="handle">
        ''' A handle to a native ITEMIDLIST.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The result of the conversion.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Widening Operator CType(handle As IntPtr) As PIDL
            Return New PIDL(handle)
        End Operator

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Combines the specified <see cref="PIDL"/> instances to create a new one.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="first">
        ''' The first <see cref="PIDL"/>.
        ''' </param>
        ''' 
        ''' <param name="second">
        ''' The second <see cref="PIDL"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A managed <see cref="PIDL"/> instance that contains both supplied lists in their respective order.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Operator +(first As PIDL, second As PIDL) As PIDL
            Return PIDL.Combine(first, second)
        End Operator

#End Region

#Region " Child Classes "

#Region " PIDLEnumerator "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Supports a simple iteration over a <see cref="IEnumerator(Of PIDL)"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private NotInheritable Class PIDLEnumerator : Implements IEnumerator(Of PIDL)

#Region " Private Fields "

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' The handle to the source ITEMIDLIST.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            Private ReadOnly pidl As IntPtr

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' The handle to the ITEMIDLIST in the collection at the current position of the enumerator.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            Private currentPidl As IntPtr

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Flag for <see cref="PIDLEnumerator.MoveNext"/>.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            Private start As Boolean

#End Region

#Region " Properties "

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Gets the element in the collection at the current position of the enumerator.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            Public ReadOnly Property Current() As PIDL Implements IEnumerator(Of PIDL).Current
                Get
                    Return If(Me.currentPidl = IntPtr.Zero, IntPtr.Zero, NativeMethods.ILCloneFirst(Me.currentPidl))
                End Get
            End Property

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Gets the element in the collection at the current position of the enumerator.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            Private ReadOnly Property IEnumerator_Current() As Object Implements IEnumerator.Current
                Get
                    Return Me.Current
                End Get
            End Property

#End Region

#Region " Constructors "

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Initializes a new instance of the <see cref="PIDLEnumerator"/> class.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <param name="handle">
            ''' A handle to the source ITEMIDLIST.
            ''' </param>
            ''' ----------------------------------------------------------------------------------------------------
            Public Sub New(handle As IntPtr)
                Me.start = True
                Me.pidl = handle
                Me.currentPidl = IntPtr.Zero
            End Sub

#End Region

#Region " Public Methods "

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Advances the enumerator to the next element of the collection.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <returns>
            ''' <see langword="True"/> if the enumerator was successfully advanced to the next element; 
            ''' <see langword="False"/> if the enumerator has passed the end of the collection.
            ''' </returns>
            ''' ----------------------------------------------------------------------------------------------------
            Public Function MoveNext() As Boolean Implements IEnumerator(Of PIDL).MoveNext
                If Me.start Then
                    Me.start = False
                    Me.currentPidl = Me.pidl
                    Return True
                End If

                Dim newPidl As IntPtr = ILNext(Me.currentPidl)
                If ILIsEmpty(newPidl) Then
                    Return False
                End If
                Me.currentPidl = newPidl
                Return True
            End Function

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Sets the enumerator to its initial position, which is before the first element in the collection.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            Public Sub Reset() Implements IEnumerator(Of PIDL).Reset
                Me.start = True
                Me.currentPidl = IntPtr.Zero
            End Sub

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            Public Sub Dispose() Implements IDisposable.Dispose
                ' Do Nothing.
            End Sub

#End Region

#Region " Private Methods "

#End Region

        End Class

#End Region

#End Region

    End Class

End Namespace

#End Region

#End If