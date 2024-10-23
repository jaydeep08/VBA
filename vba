Sub ChangeColorEverywhere()
    ' Clear previous formatting to search correctly
    Selection.Find.ClearFormatting
    Selection.Find.Font.Color = 15773696 ' Original color to replace (adjust if needed)

    ' Set the replacement formatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Color = -587137038 ' Black color

    With Selection.Find
        .Text = "" ' Match all text
        .Replacement.Text = "" ' Keep text unchanged
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With

    ' Apply the color change across the entire document
    Selection.Find.Execute Replace:=wdReplaceAll

    ' Loop through all tables and change text color inside them
    Dim tbl As Table
    Dim cell As Cell
    Dim para As Paragraph

    For Each tbl In ActiveDocument.Tables
        For Each cell In tbl.Range.Cells
            For Each para In cell.Range.Paragraphs
                para.Range.Font.Color = -587137038 ' Black color
            Next para
        Next cell
    Next tbl

    ' Also handle any content controls (if Power Automate used them)
    Dim cc As ContentControl
    For Each cc In ActiveDocument.ContentControls
        cc.Range.Font.Color = -587137038 ' Black color
    Next cc

End Sub
