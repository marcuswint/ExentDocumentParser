Imports System.Text.RegularExpressions

Module Utils_RegExContentSearch
    Public Function RegExSearch(Content As Object, SearchText As String) As String
        Dim sInput As String
        Dim sResult As String = ""
        Dim SearchPattern As String = ""
        Dim Row As Long
        Dim Column As Long

        For Row = 1 To UBound(Content, 1)
            For Column = 1 To UBound(Content, 2)

                ' /b - Text begins with
                SearchPattern = "\b" + SearchText

                If SearchPattern <> "" Then
                    sInput = Content(Row, Column)

                    If sInput <> "" Then

                        Dim regex As Regex = New Regex(SearchPattern, RegexOptions.IgnoreCase Or RegexOptions.Multiline)
                        Dim result As MatchCollection = regex.Matches(sInput)

                        If result.Count > 0 Then
                            sResult = result(0).Value
                        End If
                    End If
                End If

                ' Check if we have found result - exit loop 
                If sResult <> "" Then Exit For
            Next

            ' Check if we have found result - exit loop 
            If sResult <> "" Then Exit For
        Next

        Return sResult

    End Function

End Module
