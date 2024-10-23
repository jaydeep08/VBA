Sub ChangeTextBetweenBracketsToBlue()
    Dim tbl As Table
    Dim cell As Cell
    Dim rng As Range

    ' Loop through all tables in the document
    For Each tbl In ActiveDocument.Tables
        For Each cell In tbl.Range.Cells
            Set rng = cell.Range
            rng.MoveEnd Unit:=wdCharacter, Count:=-1 ' Exclude end-of-cell marker
            ApplyColorToTextInBrackets rng
        Next cell
    Next tbl

    ' Apply the same logic to non-table content
    Set rng = ActiveDocument.Content
    ApplyColorToTextInBrackets rng
End Sub

Sub ApplyColorToTextInBrackets(rng As Range)
    With rng.Find
        .ClearFormatting
        .Text = "\<*\>" ' Search for text inside < and >
        .Replacement.ClearFormatting
        .Replacement.Font.Color = wdColorBlue ' Set font color to blue
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True ' Use wildcard search
        .Execute Replace:=wdReplaceAll
    End With
End Sub
