Sub ChangeTextInBracketsWithControls()
    Dim tbl As Table
    Dim cell As Cell
    Dim cc As ContentControl
    Dim rng As Range

    ' Loop through all tables and cells
    For Each tbl In ActiveDocument.Tables
        For Each cell In tbl.Range.Cells
            Set rng = cell.Range
            rng.MoveEnd Unit:=wdCharacter, Count:=-1 ' Ignore end-of-cell marker
            
            ' Check for content controls within the cell
            If cell.Range.ContentControls.Count > 0 Then
                For Each cc In cell.Range.ContentControls
                    ApplyBracketFormatting cc.Range ' Apply formatting to control content
                Next cc
            Else
                ApplyBracketFormatting rng ' Apply formatting to plain cell text
            End If
        Next cell
    Next tbl

    ' Handle regular document content outside tables
    Set rng = ActiveDocument.Content
    ApplyBracketFormatting rng
End Sub

Sub ApplyBracketFormatting(rng As Range)
    With rng.Find
        .ClearFormatting
        .Text = "\<*\>" ' Wildcard to match text inside < and >
        .Replacement.ClearFormatting
        .Replacement.Font.Color = wdColorBlue ' Change color to blue
        .Forward = True
        .Wrap = wdFindStop ' Stop at the end of range
        .Format = True
        .MatchWildcards = True ' Enable wildcard matching
        .Execute Replace:=wdReplaceAll
    End With
End Sub
