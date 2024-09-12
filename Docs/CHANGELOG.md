# Easy Link File Viewer Change Log 📋

## v1.6.2 **(current)** 🆕

#### 🛠️ Fixes:
 - The shortcut's window state attribute (normal, maximized, minimized) was not being properly recognized. Solves Issue [#14](https://github.com/ElektroStudios/Easy-Link-File-Viewer/issues/14)

#### 🌟 Improvements:
 - Added keyboard shortcuts to toolbar items, thanks to @[BenVlodgi](https://github.com/ElektroStudios/Easy-Link-File-Viewer/pull/12)'s contribution.

## v1.6.1 🔄

#### 🛠️ Fixes:
 - An user settings conflict that caused an unhandled exception when trying to persist the property grid view mode.

## v1.6 🔄

#### 🚀 New Features:
 - Added experimental support to read/write Windows Installer link files. Solves Issue [#5](https://github.com/ElektroStudios/Easy-Link-File-Viewer/issues/5)
 - Added a new menu option: "Show Link Editor Toolbar". Solves Issue [#11](https://github.com/ElektroStudios/Easy-Link-File-Viewer/issues/11)
 - Added a new menu option: "Show 'Raw' tab". Solves Issue [#11](https://github.com/ElektroStudios/Easy-Link-File-Viewer/issues/11)
 - Added a new menu option: "Hide 'Recent...' files list". Solves Issue [#11](https://github.com/ElektroStudios/Easy-Link-File-Viewer/issues/11)

#### 🛠️ Fixes:
 - The program can crash when it is unable to retrieve the description field for some kind of link files.
 - The menu options "Default" and "Dark" in "Settings > Visual Style" allowed them to be unchecked individually.

#### 🌟 Improvements:
 - "About" menu button is now aligned to the top right corner of the window bounds. Solves Issue [#11](https://github.com/ElektroStudios/Easy-Link-File-Viewer/issues/11)
 - Now a message box will display to inform about errors if the data of a link file can't be read.
 - Menu item "Show toolbar" has been renamed to "Show File Menu Toolbar".

## v1.5 🔄

#### 🚀 New Features:
 - Added support for opening \*.lnk files from command-line. 
 - Added a new menu option: "Remember Window Size and Position". 
 - Added a new menu option: "Add Explorer's context-menu entry for *.lnk". Solves Issue [#2](https://github.com/ElektroStudios/Easy-Link-File-Viewer/issues/2)
 - Added an icons toolbar below the top menu strip to have faster access to the "File" menu commands. Solves Issue [#8](https://github.com/ElektroStudios/Easy-Link-File-Viewer/issues/8)

#### 🛠️ Fixes:
 - Some elements within the user interface remained unaffected when the user opted for a different font size.
 - Made capable to save changes in read-only link files. Solves Issue [#9](https://github.com/ElektroStudios/Easy-Link-File-Viewer/issues/9)
 - Fixed buffer sizes to read/write link properties that prevents exceeding the maximum allowed length. Solves Issue [#7](https://github.com/ElektroStudios/Easy-Link-File-Viewer/issues/7)

#### 🌟 Improvements:
 - Assigned a default accelerator key for all of the menu options.
 - Added a menu option to show or hide the icons toolbar.
 - Renamed / simplified names for tabs.
 - Now the user settings are stored within the executable directory in a folder with name "cache". This makes the application 100% portable. Solves Issue [#10](https://github.com/ElektroStudios/Easy-Link-File-Viewer/issues/10)
 - Added a confirmation message box for successful saves, and another for failed saves.
 - Incremented selectable font sizes up to 16pt.

## v1.4 🔄

#### 🚀 New Features:
 - Added basic support for command-line arguments. Now you can pass a shortcut's file path to the executable file to open/load the specified shortcut in the program.
 - Added a Hexadecimal Viewer (read-only) to view the raw contents of the current loaded shortcut file.

#### 🌟 Improvements:
 - Minor UI changes and source-code optimizations.

## v1.3 🔄

#### 🛠️ Fixes:
 - Fixed a visual issue with the menu strip.

## v1.2 🔄

#### 🌟 Improvements:
 - The UI editor for selecting the icon index now allows to view and select a specific icon within the resource file. (see the screenshot below)
 - The system dialogs to select files and folders now are forcefully shown in English language. Except the dialog to choose an icon index.
 - Now the application's context menu is also accessible through right click on the bottom status bar.
 - The ordering of the items in the 'File' menu has been simplified.
 - The algorithm for representing the file size of the current .lnk file has been optimized.

## v1.1 🔄

#### 🚀 New Features:
 - Added an UI editor for 'Icon' property to select a icon file through a dialog window.
 - Added an UI editor for 'Icon Index' property to preview the current icon.
 - Added an UI editor for 'Target' property to select a file or folder through a dialog window.
 - Added an UI editor for 'Working Directory' property to select a folder through a dialog window.
 - Added icon preview for the current shortcut file on the status bar.
 - Added a property 'Target Display Name' which will display a friendly name for special targets. 
      (eg. '::{20D04FE0-3AEA-1069-A2D8-08002B30309D}' CLSID is translated as 'My PC').
 - Added a context menu accessible through a right mouse click on the property grid, with the next commands:

      Open Shortcut
      Open Target
      Open Target with Arguments
      View Shortcut in Explorer
      View Target in Explorer
      View Working Directory in Explorer
      View Icon in Explorer

 - Added 'New' option in 'File' menu to create a new, empty shortcut.
 - Added 'Font Size' option in 'Settings' menu to change the UI font size.

#### 🌟 Improvements:
 - Increased default UI font size to 10pt.
 - Improved the representation of 'Length' property value.
      (Now it displays a proper file size string instead of a raw bytes-length value.)
 - Improved the resizing of the property grid.
      (Now it keeps the default column size while the user is resizing the window.)
 - Other minor UI corrections.
 - Some property descriptions were abbreviated.

## v1.0 🔄

Initial Release.