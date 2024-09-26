Private Sub Document_Open()
    ' Change the text color to blue for text between < and >
    FormatTextWithWildcard "\<*\>", RGB(0, 0, 255), False, False, 0 ' Blue Color, No Bold, No Replace, No Font Size Change

    ' Make the text bold and increase size for text between << and >>
    FormatTextWithWildcard "\<\<*\>\>", RGB(0, 0, 0), True, False, 14 ' No Color Change, Bold, No Replace, Font Size 14

    ' Hide the text and the tags between <<< and >>>
    FormatTextWithWildcard "\<\<\<*\>\>\>", RGB(0, 0, 0), False, True, 0 ' No Color Change, No Bold, Replace with Blank, No Font Size Change
End Sub

' Function to format the text using wildcards
Sub FormatTextWithWildcard(wildcardPattern As String, textColor As Long, makeBold As Boolean, replaceWithBlank As Boolean, fontSize As Integer)
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

                ' Apply color, bold, and font size if specified
                rng.Font.Color = textColor
                rng.Font.Bold = makeBold
                If fontSize > 0 Then
                    rng.Font.Size = fontSize ' Set font size if specified
                End If
            End If

            rng.Collapse wdCollapseEnd
        Loop
    End With
End Sub
