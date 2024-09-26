Private Sub Document_Open()
    ' Change the text color to blue for text between < and >
    FormatTextBetweenTags "<", ">", RGB(0, 0, 255), False, False ' Blue Color, No Bold, No Replace

    ' Make the text bold for text between << and >>
    FormatTextBetweenTags "<<", ">>", RGB(0, 0, 0), True, False ' No Color Change, Bold, No Replace

    ' Hide the text and the tags between <<< and >>>
    FormatTextBetweenTags "<<<", ">>>", RGB(0, 0, 0), False, True ' No Color Change, No Bold, Replace with Blank
End Sub

' Function to format the text between specific tags
Sub FormatTextBetweenTags(startTag As String, endTag As String, textColor As Long, makeBold As Boolean, replaceWithBlank As Boolean)
    Dim rng As Range
    Set rng = ActiveDocument.Content

    With rng.Find
        ' Search for the pattern with wildcards
        .Text = startTag & "[!<>]@" & endTag
        .MatchWildcards = True

        Do While .Execute
            ' Set the range to modify the text
            Set rng = rng.Duplicate

            ' Expand the range to include the tags
            rng.MoveStart wdCharacter, 0
            rng.MoveEnd wdCharacter, 0

            If replaceWithBlank Then
                ' Replace both the tags and the content inside with a blank
                rng.Text = ""
            Else
                ' Apply formatting if not replacing the text
                rng.MoveStart wdCharacter, Len(startTag)
                rng.MoveEnd wdCharacter, -Len(endTag)
                rng.Font.Color = textColor
                rng.Font.Bold = makeBold
            End If

            rng.Collapse wdCollapseEnd
        Loop
    End With
End Sub
