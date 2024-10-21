Sub FormatAndCleanText()
    Dim para As Paragraph
    Dim rng As Range, tempRng As Range, innerText As Range

    ' Define color codes
    Dim targetColor As Long: targetColor = RGB(255, 0, 0) ' Original color to be replaced
    Dim blackColor As Long: blackColor = RGB(0, 0, 0) ' Black (new color for paragraphs)
    Dim blueColor As Long: blueColor = RGB(0, 0, 255) ' Blue (for text inside <...>)

    ' Disable screen updates for performance
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Step 1: Format and remove <heading> tags
    For Each para In ActiveDocument.Paragraphs
        Set rng = para.Range
        With rng.Find
            .ClearFormatting
            .Text = "<heading>*</heading>"
            .MatchWildcards = True
            .Wrap = wdFindContinue ' Ensure search continues across the document
            Do While .Execute
                Set tempRng = rng.Duplicate
                tempRng.Start = tempRng.Start + 9 ' Skip <heading>
                tempRng.End = tempRng.End - 10 ' Exclude </heading>

                tempRng.Font.Color = blueColor
                tempRng.Font.Bold = True

                ' Remove <heading> tags by keeping only inner text
                rng.Text = tempRng.Text
                rng.Collapse Direction:=wdCollapseEnd
            Loop
        End With
    Next para

    ' Step 2: Change paragraphs with specific color to black
    For Each para In ActiveDocument.Paragraphs
        Set rng = para.Range
        If rng.Font.Color = targetColor Then
            rng.Font.Color = blackColor
        End If
    Next para

    ' Step 3: Make text inside <...> blue
    For Each para In ActiveDocument.Paragraphs
        Set rng = para.Range
        With rng.Find
            .ClearFormatting
            .Text = "<*>"
            .MatchWildcards = True
            .Wrap = wdFindContinue ' Continue search across the document
            Do While .Execute
                Set innerText = rng.Duplicate
                innerText.Font.Color = blueColor
                rng.Collapse Direction:=wdCollapseEnd
            Loop
        End With
    Next para

    ' Re-enable screen updates
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Formatting complete!"
End Sub
