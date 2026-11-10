#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports Microsoft.VisualStudio.TestTools.UnitTesting

#End Region

#Region " Assembly Attributes "

' Several tests redirect the process-wide Console.Out / Console.Error streams
' and create real shortcut files through the Windows Shell COM objects.
' Running them sequentially keeps those shared resources deterministic.
<Assembly: DoNotParallelize>

#End Region
