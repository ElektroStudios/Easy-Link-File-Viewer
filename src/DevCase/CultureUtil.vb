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

Imports System.Collections.ObjectModel
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Security

Imports DevCase.Interop.Unmanaged.Win32
Imports DevCase.Interop.Unmanaged.Win32.Enums

#End Region

#Region " Culture Util "

Namespace DevCase.Core.Application.Tools

    Friend NotInheritable Class CultureUtil

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sets the preferred UI languages for the current process.
        ''' <para></para>
        ''' To retrieve the preferred UI languages for the current process, 
        ''' call <see cref="GetProcessPreferredUILanguages()"/> function.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Dim langNames As String() = {"en-US", "es-ES", "it-IT", "de-DE", "fr-FR"}
        ''' Dim successfulLangs As Integer = SetProcessPreferredUILanguages(langNames)
        ''' Console.WriteLine($"{NameOf(successfulLangs)}: {successfulLangs}")
        ''' 
        ''' Dim currentPreferredLangs As ReadOnlyCollection(Of CultureInfo) = GetProcessPreferredUILanguages()
        ''' Console.WriteLine($"{NameOf(currentPreferredLangs)}: {String.Join(", ", currentPreferredLangs.Select(Function(ci) ci.Name))}")
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="langNames">
        ''' The name of the languages, in the format languagecode2-country/regioncode2 (eg. "en-US")
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' Returns the amount of languages that were successfully set.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function SetProcessPreferredUILanguages(ParamArray langNames As String()) As Integer

            If (langNames Is Nothing) Then
                Throw New ArgumentNullException(paramName:=NameOf(langNames))
            End If

            Dim languages As String = String.Join(ControlChars.NullChar, langNames)
            Dim numLangs As UInteger
            Dim result As Boolean = NativeMethods.SetProcessPreferredUILanguages(MuiLanguageMode.Name, languages, numLangs)
            If Not (result) Then
                Return 0
            End If

            Return CType(numLangs, Integer)

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sets the preferred UI languages for the current process.
        ''' <para></para>
        ''' To retrieve the preferred UI languages for the current process, 
        ''' call <see cref="GetProcessPreferredUILanguages()"/> function.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Dim cultures As CultureInfo() = {
        '''     CultureInfo.CreateSpecificCulture("en-US"),
        '''     CultureInfo.CreateSpecificCulture("es-ES"),
        '''     CultureInfo.CreateSpecificCulture("it-IT"),
        '''     CultureInfo.CreateSpecificCulture("de-DE"),
        '''     CultureInfo.CreateSpecificCulture("fr-FR")
        ''' }
        ''' 
        ''' Dim successfulLangs As Integer = SetProcessPreferredUILanguages(cultures)
        ''' Console.WriteLine($"{NameOf(successfulLangs)}: {successfulLangs}")
        ''' 
        ''' Dim currentPreferredLangs As ReadOnlyCollection(Of CultureInfo) = GetProcessPreferredUILanguages()
        ''' Console.WriteLine($"{NameOf(currentPreferredLangs)}: {String.Join(", ", currentPreferredLangs.Select(Function(ci) ci.Name))}")
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="cultures">
        ''' An array of <see cref="CultureInfo"/> representing each language.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' Returns the amount of languages that were successfully set.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function SetProcessPreferredUILanguages(ParamArray cultures As CultureInfo()) As Integer
            Return SetProcessPreferredUILanguages(cultures.Select(Function(ci) ci.Name).ToArray())
        End Function

#End Region

    End Class

End Namespace

#End Region
