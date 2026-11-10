#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports Microsoft.VisualStudio.TestTools.UnitTesting

#End Region

#Region " TextAssert "

Namespace EasyLink.Tests.Helpers

    ''' <summary>
    ''' Provides string assertions with descriptive failure messages.
    ''' </summary>
    Friend NotInheritable Class TextAssert

#Region " Constructors "

        ''' <summary>
        ''' Prevents a default instance of the <see cref="TextAssert"/> class from being created.
        ''' </summary>
        Private Sub New()
        End Sub

#End Region

#Region " Public Methods "

        ''' <summary>
        ''' Asserts that <paramref name="actual"/> contains <paramref name="expectedSubstring"/>
        ''' using an ordinal, case-sensitive comparison.
        ''' </summary>
        '''
        ''' <param name="actual">
        ''' The text to search in.
        ''' </param>
        '''
        ''' <param name="expectedSubstring">
        ''' The substring that must be present.
        ''' </param>
        Friend Shared Sub Contains(actual As String, expectedSubstring As String)

            If expectedSubstring Is Nothing Then
                Throw New ArgumentNullException(NameOf(expectedSubstring))
            End If

            Assert.IsNotNull(actual, $"Expected text containing ""{expectedSubstring}"", but the actual text is null.")
            Assert.IsTrue(actual.IndexOf(expectedSubstring, StringComparison.Ordinal) >= 0,
                          $"Expected text containing ""{expectedSubstring}"". Actual text: ""{actual}""")

        End Sub

        ''' <summary>
        ''' Asserts that <paramref name="actual"/> does not contain <paramref name="unexpectedSubstring"/>
        ''' using an ordinal, case-sensitive comparison.
        ''' </summary>
        '''
        ''' <param name="actual">
        ''' The text to search in.
        ''' </param>
        '''
        ''' <param name="unexpectedSubstring">
        ''' The substring that must not be present.
        ''' </param>
        Friend Shared Sub DoesNotContain(actual As String, unexpectedSubstring As String)

            If unexpectedSubstring Is Nothing Then
                Throw New ArgumentNullException(NameOf(unexpectedSubstring))
            End If

            Assert.IsNotNull(actual, "The actual text is null.")
            Assert.IsTrue(actual.IndexOf(unexpectedSubstring, StringComparison.Ordinal) < 0,
                          $"Expected text not containing ""{unexpectedSubstring}"". Actual text: ""{actual}""")

        End Sub

        ''' <summary>
        ''' Asserts that two file system paths are equal, ignoring case and expanding
        ''' any environment variable that the Windows Shell may have stored in place of a path prefix.
        ''' </summary>
        '''
        ''' <param name="expectedPath">
        ''' The expected path.
        ''' </param>
        '''
        ''' <param name="actualPath">
        ''' The actual path.
        ''' </param>
        Friend Shared Sub PathEquals(expectedPath As String, actualPath As String)

            Assert.IsNotNull(actualPath, $"Expected path ""{expectedPath}"", but the actual path is null.")

            Dim expandedActual As String = Environment.ExpandEnvironmentVariables(actualPath)
            Assert.IsTrue(String.Equals(expectedPath, expandedActual, StringComparison.OrdinalIgnoreCase),
                          $"Expected path ""{expectedPath}"". Actual path: ""{actualPath}"" (expanded: ""{expandedActual}"")")

        End Sub

#End Region

    End Class

End Namespace

#End Region
