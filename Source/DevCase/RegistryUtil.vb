Imports Microsoft.Win32

Module RegistryUtil

    ''' <summary>
    ''' Creates a registry key that represents a new entry in the Explorer's context-menu for the specified file type.
    ''' </summary>
    ''' 
    ''' <param name="fileType">
    ''' The file type (typically a file extension) for which to create the entry in the Explorer's context-menu.
    ''' </param>
    ''' 
    ''' <param name="keyName">
    ''' The name of the registry key.
    ''' </param>
    ''' 
    ''' <param name="text">
    ''' The display text for the entry in the Explorer's context-menu.
    ''' <para></para>
    ''' This value can be null, in which case <paramref name="keyName"/> will be used as text.
    ''' </param>
    ''' 
    ''' <param name="position">
    ''' The position of the entry in the Explorer's context-menu.
    ''' <para></para>
    ''' Valid values are: "top", "middle" and "bottom".
    ''' <para></para>
    ''' This value can be null.
    ''' </param>
    ''' 
    ''' <param name="icon">
    ''' The icon to show for the entry in the Explorer's context-menu.
    ''' <para></para>
    ''' This value can be null.
    ''' </param>
    ''' 
    ''' <param name="command">
    ''' The command to execute when the entry is clicked in the Explorer's context-menu.
    ''' </param>
    <DebuggerStepThrough>
    Public Sub CreateFileTypeRegistryMenuEntry(fileType As String, keyName As String,
                                               text As String, position As String,
                                               icon As String, command As String)

        If String.IsNullOrWhiteSpace(fileType) Then
            Throw New ArgumentNullException(paramName:=NameOf(fileType))
        End If

        If String.IsNullOrWhiteSpace(keyName) Then
            Throw New ArgumentNullException(paramName:=NameOf(keyName))
        End If

        If String.IsNullOrWhiteSpace(command) Then
            Throw New ArgumentNullException(paramName:=NameOf(command))
        End If

        If String.IsNullOrEmpty(text) Then
            text = keyName
        End If

        Using rootKey As RegistryKey = Registry.ClassesRoot,
              subKey As RegistryKey = rootKey.CreateSubKey($"{fileType}\shell\{keyName}", writable:=True),
              subKeyCommand As RegistryKey = subKey.CreateSubKey("command", writable:=True)

            subKey.SetValue("", text, RegistryValueKind.String)
            subKey.SetValue("icon", icon, RegistryValueKind.String)
            subKey.SetValue("position", position, RegistryValueKind.String)

            subKeyCommand.SetValue("", command, RegistryValueKind.String)
        End Using

    End Sub

    ''' <summary>
    ''' Deletes an existing registry key representing an entry in the Explorer's context-menu for the specified file type.
    ''' </summary>
    ''' 
    ''' <param name="fileType">
    ''' The file type associated with the registry entry.
    ''' </param>
    ''' 
    ''' <param name="keyName">
    ''' The name of the registry key to delete.
    ''' </param>
    ''' 
    ''' <param name="throwOnMissingsubKey">
    ''' Optional. If <see langword="True"/>, throws an exception if the registry key is not found.
    ''' <para></para>
    ''' Default value is <see langword="True"/>.
    ''' </param>
    <DebuggerStepThrough>
    Public Sub DeleteFileTypeRegistryMenuEntry(fileType As String, keyName As String, Optional throwOnMissingsubKey As Boolean = True)

        If String.IsNullOrWhiteSpace(fileType) Then
            Throw New ArgumentNullException(paramName:=NameOf(fileType))
        End If

        If String.IsNullOrWhiteSpace(keyName) Then
            Throw New ArgumentNullException(paramName:=NameOf(keyName))
        End If

        Using rootKey As RegistryKey = Registry.ClassesRoot

            rootKey.DeleteSubKeyTree($"{fileType}\shell\{keyName}", throwOnMissingsubKey)
        End Using

    End Sub

End Module
