# VBA

Private Sub Document_Open()
    ' Change the color of the text between < and >
    ChangeTextColorBetweenTags "<", ">", RGB(255, 0, 0) 'Red Color

    ' Bold and change the color of text between <header> and </header>
    FormatTextBetweenTags "<header>", "</header>", RGB(0, 0, 255), True 'Blue Color and Bold

    ' Hide the text between <height> and </height>
    HideTextBetweenTags "<height>", "</height>"
End Sub

Sub ChangeTextColorBetweenTags(startTag As String, endTag As String, textColor As Long)
    Dim rng As Range
    Set rng = ActiveDocument.Content

    With rng.Find
        .Text = startTag & "*" & endTag
        .MatchWildcards = True

        Do While .Execute
            rng.Font.Color = textColor
            rng.Collapse wdCollapseEnd
        Loop
    End With
End Sub

Sub FormatTextBetweenTags(startTag As String, endTag As String, textColor As Long, makeBold As Boolean)
    Dim rng As Range
    Set rng = ActiveDocument.Content

    With rng.Find
        .Text = startTag & "*" & endTag
        .MatchWildcards = True

        Do While .Execute
            ' Remove the start and end tags
            rng.Text = Replace(rng.Text, startTag, "")
            rng.Text = Replace(rng.Text, endTag, "")
            
            ' Apply formatting
            rng.Font.Color = textColor
            rng.Font.Bold = makeBold
            rng.Collapse wdCollapseEnd
        Loop
    End With
End Sub

Sub HideTextBetweenTags(startTag As String, endTag As String)
    Dim rng As Range
    Set rng = ActiveDocument.Content

    With rng.Find
        .Text = startTag & "*" & endTag
        .MatchWildcards = True

        Do While .Execute
            ' Remove the start and end tags
            rng.Text = Replace(rng.Text, startTag, "")
            rng.Text = Replace(rng.Text, endTag, "")
            
            ' Hide the text
            rng.Font.Hidden = True
            rng.Collapse wdCollapseEnd
        Loop
    End With
End Sub
