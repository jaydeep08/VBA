Sub ChangeColorForBracesText()
    Dim rng As Range
    Set rng = ActiveDocument.Content

    With rng.Find
        .Text = "<(*)>"
        .MatchWildcards = True  ' Use wildcard search to find text inside angular braces
        .Execute
    End With

    Do While rng.Find.Found
        rng.Font.Color = wdColorBlue  ' Change this color to your preference
        rng.Collapse wdCollapseEnd    ' Move to the next occurrence
        rng.Find.Execute
    Loop
End Sub
