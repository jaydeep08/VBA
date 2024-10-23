Sub ChangeTextBetweenBracketsToBlue()
    Dim tbl As Table
    Dim cell As Cell
    Dim rng As Range

    ' Set the color to blue
    Const BLUE_COLOR As Long = RGB(0, 0, 255)

    ' Loop through the entire document, including tables
    For Each tbl In ActiveDocument.Tables
        For Each cell In tbl.Range.Cells
            Set rng = cell.Range
            rng.MoveEnd Unit:=wdCharacter, Count:=-1 ' Exclude end-of-cell marker
            ApplyColorToTextInBrackets rng, BLUE_COLOR
        Next cell
    Next tbl

    ' Also apply the same logic to non-table text
    Set rng = ActiveDocument.Content
    ApplyColorToTextInBrackets rng, BLUE_COLOR
End Sub

Sub ApplyColorToTextInBrackets(rng As Range, color As Long)
    With rng.Find
        .ClearFormatting
        .Text = "\<*\>" ' Search pattern for text inside < and >
        .Replacement.ClearFormatting
        .Replacement.Font.Color = color
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True ' Enable wildcard search
        .Execute Replace:=wdReplaceAll
    End With
End Sub
