Sub ChangeColorInAngularBraces()
    Dim doc As Document
    Dim rng As Range

    ' Set the document and range to search the entire document
    Set doc = ActiveDocument
    Set rng = doc.Content

    ' Configure the search properties
    With rng.Find
        .Text = "<(*)>"
        .MatchWildcards = True  ' Use wildcards to match text inside angular braces
        .Wrap = wdFindStop      ' Stop searching at the end of the document
    End With

    ' Loop through all matches
    Do While rng.Find.Execute = True
        ' Check if the match is inside a content control
        If rng.ContentControls.Count > 0 Then
            Dim cc As ContentControl
            For Each cc In rng.ContentControls
                ' Apply color to the text inside the content control
                cc.Range.Font.Color = wdColorBlue  ' Set to your preferred color
            Next cc
        Else
            ' Apply color to the text outside of content controls
            rng.Font.Color = wdColorBlue  ' Set to your preferred color
        End If

        ' Move the range to avoid infinite loop
        rng.Collapse wdCollapseEnd  ' Move the range to the end of the current match
    Loop
End Sub
