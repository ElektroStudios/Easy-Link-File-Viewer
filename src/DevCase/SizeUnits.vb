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

#Region " Size Units "

Namespace DevCase.Core.IO

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Specifies a size unit.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Friend Enum SizeUnits As Long

        ''' <summary>
        ''' Represents 1 Byte.
        ''' <para></para>
        ''' (8 Bits)
        ''' </summary>
        [Byte] = CLng(2 ^ 0)

        ''' <summary>
        ''' Represents 1 Kilobyte.
        ''' <para></para>
        ''' (1.024 Bytes)
        ''' </summary>
        Kilobyte = CLng(2 ^ 10)

        ''' <summary>
        ''' Represents 1 MegaByte.
        ''' <para></para>
        ''' (1.048.576 Bytes)
        ''' </summary>
        Megabyte = CLng(2 ^ 20)

        ''' <summary>
        ''' Represents 1 Gigabyte.
        ''' <para></para>
        ''' (1.073.741.824 Bytes)
        ''' </summary>
        Gigabyte = CLng(2 ^ 30)

        ''' <summary>
        ''' Represents 1 Terabyte.
        ''' <para></para>
        ''' (1.099.511.627.776 Bytes)
        ''' </summary>
        Terabyte = CLng(2 ^ 40)

        ''' <summary>
        ''' Represents 1 Petabyte.
        ''' <para></para>
        ''' (1.125.899.906.842.624 Bytes)
        ''' </summary>
        Petabyte = CLng(2 ^ 50)

        ''' <summary>
        ''' Represents 1 Exabyte.
        ''' <para></para>
        ''' (1.152.921.504.606.846.976 Bytes)
        ''' </summary>
        Exabyte = CLng(2 ^ 60)

    End Enum

End Namespace

#End Region
