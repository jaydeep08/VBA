Private Sub Document_Open()
    ' Change the text color to blue for text between < and >
    FormatTextWithWildcard "\<*\>", RGB(0, 0, 255), False, False ' Blue Color, No Bold, No Replace

    ' Make the text bold for text between << and >>
    FormatTextWithWildcard "\<\<*\>\>", RGB(0, 0, 0), True, False ' No Color Change, Bold, No Replace

    ' Hide the text and the tags between <<< and >>>
    FormatTextWithWildcard "\<\<\<*\>\>\>", RGB(0, 0, 0), False, True ' No Color Change, No Bold, Replace with Blank
End Sub

' Function to format the text using wildcards
Sub FormatTextWithWildcard(wildcardPattern As String, textColor As Long, makeBold As Boolean, replaceWithBlank As Boolean)
    Dim rng As Range
    Set rng = ActiveDocument.Content

    With rng.Find
        ' Use wildcard search
        .Text = wildcardPattern
        .MatchWildcards = True

        Do While .Execute
            ' Set the range to the found text
            Set rng = rng.Duplicate

            If replaceWithBlank Then
                ' Replace the entire found text (tags and content) with a blank
                rng.Text = ""
            Else
                ' Apply formatting to the text inside the tags
                rng.MoveStart wdCharacter, 1 ' Move past the opening tag
                rng.MoveEnd wdCharacter, -1 ' Move before the closing tag

                ' Apply color and bold
                rng.Font.Color = textColor
                rng.Font.Bold = makeBold
            End If

            rng.Collapse wdCollapseEnd
        Loop
    End With
End Sub
