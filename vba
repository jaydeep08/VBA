Sub FormatTextBasedOnRules()
    Dim para As Paragraph
    Dim rng As Range
    Dim startPos As Long, endPos As Long

    ' Define your color codes here
    Dim targetColor As Long: targetColor = RGB(255, 0, 0) ' Example: Red (old color)
    Dim newColor As Long: newColor = RGB(0, 128, 0) ' Example: Green (new color)
    Dim blueColor As Long: blueColor = RGB(0, 0, 255) ' Example: Blue
    Dim headingColor As Long: headingColor = RGB(128, 0, 128) ' Example: Purple

    ' Iterate through all paragraphs
    For Each para In ActiveDocument.Paragraphs
        Set rng = para.Range

        ' 1. Change color of specific text (targetColor to newColor)
        If rng.Font.Color = targetColor Then
            rng.Font.Color = newColor
        End If

        ' 2. Change text between < and > to blue
        startPos = InStr(rng.Text, "<")
        Do While startPos > 0
            endPos = InStr(startPos + 1, rng.Text, ">")
            If endPos > startPos Then
                Dim innerText As Range
                Set innerText = rng.Duplicate
                innerText.Start = rng.Start + startPos
                innerText.End = rng.Start + endPos

                innerText.Font.Color = blueColor
            End If
            startPos = InStr(endPos + 1, rng.Text, "<")
        Loop

        ' 3. Make text inside <heading>...</heading> bold and change color to headingColor
        startPos = InStr(rng.Text, "<heading>")
        Do While startPos > 0
            endPos = InStr(startPos + 9, rng.Text, "</heading>")
            If endPos > startPos Then
                Dim headingText As Range
                Set headingText = rng.Duplicate
                headingText.Start = rng.Start + startPos + 9 ' After <heading>
                headingText.End = rng.Start + endPos

                headingText.Font.Bold = True
                headingText.Font.Color = headingColor
            End If
            startPos = InStr(endPos + 10, rng.Text, "<heading>")
        Loop
    Next para

    MsgBox "Formatting complete!"
End Sub
