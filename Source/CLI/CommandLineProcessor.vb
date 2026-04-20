#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.Collections.Generic
Imports System.IO
Imports System.Windows.Forms

Imports DevCase.Core.IO

#End Region

#Region " CommandLineProcessor "

''' 
''' <summary>
''' Executes the command-line operation specified by a <see cref="CommandLineOptions"/> instance.
''' <para></para>
''' Handles console attachment for output, shortcut file creation/modification/reading,
''' and exit code management.
''' </summary>
''' 
Friend NotInheritable Class CommandLineProcessor

#Region " Public Methods "

    ''' 
    ''' <summary>
    ''' Parses the specified command-line arguments and executes the requested action.
    ''' </summary>
    ''' 
    ''' <param name="args">
    ''' The command-line arguments to process.
    ''' </param>
    ''' 
    ''' <returns>
    ''' An exit code: 0 for success, 1 for argument/validation errors, 2 for runtime errors.
    ''' </returns>
    ''' 
    Friend Function Run(args As IList(Of String)) As Integer

        Try
            Dim options As CommandLineOptions = CommandLineParser.Parse(args)

            Select Case options.Action

                Case CommandLineAction.Help
                    Me.PrintHelp()
                    Return 0

                Case CommandLineAction.Create
                    Return Me.ExecuteCreate(options)

                Case CommandLineAction.Modify
                    Return Me.ExecuteModify(options)

                Case CommandLineAction.Read
                    Return Me.ExecuteRead(options)

                Case Else
                    Console.Error.WriteLine("Error: No action specified. Use /Help for usage information.")
                    Return 1

            End Select

        Catch argEx As ArgumentException
            Console.Error.WriteLine($"Argument error: {argEx.Message}")
            Return 1

        Catch ex As Exception
            Console.Error.WriteLine($"Error: {ex.Message}")
            Return 2

        End Try

    End Function

#End Region

#Region " Private Methods "

    ''' 
    ''' <summary>
    ''' Validates that the file path is not empty and has a .lnk extension.
    ''' </summary>
    ''' 
    ''' <param name="options">
    ''' The parsed command-line options.
    ''' </param>
    ''' 
    ''' <param name="actionName">
    ''' The action name for error messages (e.g. "Create", "Modify", "Read").
    ''' </param>
    ''' 
    ''' <returns>
    ''' An exit code: 0 if valid, 1 if validation failed.
    ''' </returns>
    ''' 
    Private Function ValidateFilePath(options As CommandLineOptions, actionName As String) As Integer

        If String.IsNullOrWhiteSpace(options.FilePath) Then
            Console.Error.WriteLine($"Error: /File is required for /{actionName}.")
            Return 1
        End If

        If Not options.FilePath.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase) Then
            Console.Error.WriteLine("Error: File path must have a .lnk extension.")
            Return 1
        End If

        Return 0

    End Function

    ''' 
    ''' <summary>
    ''' Executes the <see cref="CommandLineAction.Create"/> action.
    ''' Creates a new shortcut (.lnk) file with the specified properties.
    ''' </summary>
    ''' 
    ''' <param name="options">
    ''' The parsed command-line options.
    ''' </param>
    ''' 
    ''' <returns>
    ''' An exit code: 0 for success, 1 for validation errors, 2 for runtime errors.
    ''' </returns>
    ''' 
    Private Function ExecuteCreate(options As CommandLineOptions) As Integer

        Dim validationResult As Integer = Me.ValidateFilePath(options, "Create")
        If validationResult <> 0 Then
            Return validationResult
        End If

        If File.Exists(options.FilePath) AndAlso Not options.Force Then
            Console.Error.WriteLine($"Error: File already exists: {options.FilePath}")
            Console.Error.WriteLine("Use /Force to overwrite.")
            Return 2
        End If

        ' Create the ShortcutFileInfo instance in ViewMode so property
        ' assignments are held in memory until Create() persists them all.
        Try
            Dim shortcut As New ShortcutFileInfo(options.FilePath) With {.ViewMode = True}
            Me.ApplyOptionsToShortcut(shortcut, options)

            shortcut.Create()

            Console.WriteLine($"Shortcut created: {shortcut.FullName}")
            Return 0

        Catch ex As Exception
            Console.Error.WriteLine($"Error creating shortcut (HRESULT 0x{ex.HResult:X}): {ex.Message}")
            Return 2
        End Try

    End Function

    ''' 
    ''' <summary>
    ''' Executes the <see cref="CommandLineAction.Modify"/> action.
    ''' Modifies an existing shortcut (.lnk) file, changing only the specified properties.
    ''' </summary>
    ''' 
    ''' <param name="options">
    ''' The parsed command-line options.
    ''' </param>
    ''' 
    ''' <returns>
    ''' An exit code: 0 for success, 1 for validation errors, 2 for runtime errors.
    ''' </returns>
    ''' 
    Private Function ExecuteModify(options As CommandLineOptions) As Integer

        Dim validationResult As Integer = Me.ValidateFilePath(options, "Modify")
        If validationResult <> 0 Then
            Return validationResult
        End If

        If Not File.Exists(options.FilePath) Then
            Console.Error.WriteLine($"Error: File not found: {options.FilePath}")
            Return 2
        End If

        Try
            ' Constructor reads existing properties via ReadLink().
            ' ViewMode = True ensures property assignments don't write immediately.
            Dim shortcut As New ShortcutFileInfo(options.FilePath) With {.ViewMode = True}

            Me.ApplyOptionsToShortcut(shortcut, options)

            ' Create() persists all properties (original + modified) to the file.
            shortcut.Create()

            Console.WriteLine($"Shortcut modified: {shortcut.FullName}")
            Return 0

        Catch ex As Exception
            Console.Error.WriteLine($"Error modifying shortcut (HRESULT 0x{ex.HResult:X}): {ex.Message}")
            Return 2
        End Try

    End Function

    ''' 
    ''' <summary>
    ''' Executes the <see cref="CommandLineAction.Read"/> action.
    ''' Reads and displays all properties of a shortcut (.lnk) file.
    ''' </summary>
    ''' 
    ''' <param name="options">
    ''' The parsed command-line options.
    ''' </param>
    ''' 
    ''' <returns>
    ''' An exit code: 0 for success, 1 for validation errors, 2 for runtime errors.
    ''' </returns>
    ''' 
    Private Function ExecuteRead(options As CommandLineOptions) As Integer

        Dim validationResult As Integer = Me.ValidateFilePath(options, "Read")
        If validationResult <> 0 Then
            Return validationResult
        End If

        If Not File.Exists(options.FilePath) Then
            Console.Error.WriteLine($"Error: File not found: {options.FilePath}")
            Return 2
        End If

        Try
            Dim shortcut As New ShortcutFileInfo(options.FilePath) With {.ViewMode = True}

            Console.WriteLine()
            Console.WriteLine("Shortcut Properties")
            Console.WriteLine("====================")
            Console.WriteLine($"  File Name        : {shortcut.Name}")
            Console.WriteLine($"  Full Path        : {shortcut.FullName}")
            Console.WriteLine($"  Target           : {shortcut.Target}")
            Console.WriteLine($"  Target Arguments : {shortcut.TargetArguments}")
            Console.WriteLine($"  Description      : {shortcut.Description}")
            Console.WriteLine($"  Icon             : {shortcut.Icon}")
            Console.WriteLine($"  Icon Index       : {shortcut.IconIndex}")
            Console.WriteLine($"  Hotkey           : {CommandLineProcessor.FormatHotkey(shortcut.Hotkey)}")
            Console.WriteLine($"  Window State     : {shortcut.WindowState}")
            Console.WriteLine($"  Working Directory: {shortcut.WorkingDirectory}")
            Console.WriteLine($"  App Model ID     : {shortcut.AppId}")
            Console.WriteLine()

            Return 0

        Catch ex As Exception
            Console.Error.WriteLine($"Error reading shortcut (HRESULT 0x{ex.HResult:X}): {ex.Message}")
            Return 2
        End Try

    End Function

    ''' 
    ''' <summary>
    ''' Applies the specified command-line options to a <see cref="ShortcutFileInfo"/> instance.
    ''' Only properties explicitly specified on the command line are modified.
    ''' </summary>
    ''' 
    ''' <param name="shortcut">
    ''' The <see cref="ShortcutFileInfo"/> instance to modify.
    ''' </param>
    ''' 
    ''' <param name="options">
    ''' The parsed command-line options containing the property values to apply.
    ''' </param>
    ''' 
    Private Sub ApplyOptionsToShortcut(shortcut As ShortcutFileInfo, options As CommandLineOptions)

        If options.Target IsNot Nothing Then
            shortcut.Target = options.Target
        End If

        If options.TargetArguments IsNot Nothing Then
            shortcut.TargetArguments = options.TargetArguments
        End If

        If options.Description IsNot Nothing Then
            shortcut.Description = options.Description
        End If

        If options.Icon IsNot Nothing Then
            shortcut.Icon = options.Icon
        End If

        If options.IconIndex.HasValue Then
            shortcut.IconIndex = options.IconIndex.Value
        End If

        If options.Hotkey.HasValue Then
            shortcut.Hotkey = options.Hotkey.Value
        End If

        If options.WindowState.HasValue Then
            shortcut.WindowState = options.WindowState.Value
        End If

        If options.WorkingDirectory IsNot Nothing Then
            shortcut.WorkingDirectory = options.WorkingDirectory
        End If

        If options.AppId IsNot Nothing Then
            shortcut.AppId = options.AppId
        End If

    End Sub

    ''' 
    ''' <summary>
    ''' Formats a <see cref="Keys"/> value into a human-readable string for console output.
    ''' </summary>
    ''' 
    ''' <param name="hotkey">
    ''' The hotkey value to format.
    ''' </param>
    ''' 
    ''' <returns>
    ''' A formatted string representation of the hotkey (e.g. "Control + Alt + F1"),
    ''' or "None" if no hotkey is assigned.
    ''' </returns>
    ''' 
    Private Shared Function FormatHotkey(hotkey As Keys) As String

        If hotkey = Keys.None Then
            Return "None"
        End If

        ' Keys.ToString() produces "F1, Control, Alt" — we reformat it for clarity.
        Return hotkey.ToString().Replace(",", " +")

    End Function

    ''' 
    ''' <summary>
    ''' Prints the command-line usage help text to the console.
    ''' </summary>
    ''' 
    Private Sub PrintHelp()

        Dim exeName As String = Path.GetFileName(Application.ExecutablePath)

        Console.WriteLine()
        Console.WriteLine($"{My.Application.Info.ProductName} - Command Line Interface")
        Console.WriteLine()
        Console.WriteLine("Usage:")
        Console.WriteLine($"  ""{exeName}"" /Create /File:""path.lnk"" [options]")
        Console.WriteLine($"  ""{exeName}"" /Modify /File:""path.lnk"" [options]")
        Console.WriteLine($"  ""{exeName}"" /Read   /File:""path.lnk""")
        Console.WriteLine($"  ""{exeName}"" /Help")
        Console.WriteLine()
        Console.WriteLine("Actions:")
        Console.WriteLine("  /Create          Create a new shortcut file.")
        Console.WriteLine("  /Modify          Modify an existing shortcut file (only specified properties are changed).")
        Console.WriteLine("  /Read            Display all properties of a shortcut file.")
        Console.WriteLine("  /Help            Display this help text.")
        Console.WriteLine()
        Console.WriteLine("Options:")
        Console.WriteLine("  /File:""path""          Path of the shortcut (.lnk) file. (Required)")
        Console.WriteLine("  /Target:""path""        Target file or directory path.")
        Console.WriteLine("  /Args:""arguments""     Target command-line arguments.")
        Console.WriteLine("  /Desc:""text""          Shortcut description.")
        Console.WriteLine("  /Icon:""path""          Icon file path (e.g. ""Shell32.dll"").")
        Console.WriteLine("  /IconIndex:N          Icon index within the icon file (e.g. 0).")
        Console.WriteLine("  /Hotkey:keys          Hotkey combination (e.g. Control+Alt+F1, None, or numeric value).")
        Console.WriteLine("  /WindowState:state    Window state: Normal, Minimized (Min), Maximized (Max).")
        Console.WriteLine("  /WorkDir:""path""       Working directory path.")
        Console.WriteLine("  /AppId:""id""           Application User Model ID.")
        Console.WriteLine("  /Force                Overwrite existing file when using /Create.")
        Console.WriteLine()
        Console.WriteLine("Examples:")
        Console.WriteLine()
        Console.WriteLine("  Create a shortcut to Notepad:")
        Console.WriteLine($"    ""{exeName}"" /Create /File:""C:\MyShortcut.lnk"" /Target:""C:\Windows\notepad.exe"" /Desc:""Open Notepad""")
        Console.WriteLine()
        Console.WriteLine("  Modify the target of an existing shortcut:")
        Console.WriteLine($"    ""{exeName}"" /Modify /File:""C:\MyShortcut.lnk"" /Target:""C:\Windows\calc.exe""")
        Console.WriteLine()
        Console.WriteLine("  Read shortcut properties:")
        Console.WriteLine($"    ""{exeName}"" /Read /File:""C:\MyShortcut.lnk""")
        Console.WriteLine()
        Console.WriteLine("Exit Codes:")
        Console.WriteLine("  0  Success.")
        Console.WriteLine("  1  Invalid arguments or validation error.")
        Console.WriteLine("  2  Runtime error (file not found, I/O error, etc.).")
        Console.WriteLine()

    End Sub

#End Region

End Class

#End Region