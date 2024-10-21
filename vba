Sub FormatTextBasedOnRules()
    Dim para As Paragraph
    Dim rng As Range
    Dim startPos As Long, endPos As Long
    Dim findText As String

    ' Iterate through all paragraphs
    For Each para In ActiveDocument.Paragraphs
        Set rng = para.Range
        
        ' 1. Change color of specific text (e.g., red to green)
        If rng.Font.Color = wdColorRed Then
            rng.Font.Color = wdColorGreen
        End If
        
        ' 2. Change text between < and > to blue
        startPos = InStr(rng.Text, "<")
        Do While startPos > 0
            endPos = InStr(startPos + 1, rng.Text, ">")
            If endPos > startPos Then
                Set innerText = rng.Duplicate
                innerText.Start = rng.Start + startPos
                innerText.End = rng.Start + endPos
                
                innerText.Font.Color = wdColorBlue
            End If
            startPos = InStr(endPos + 1, rng.Text, "<")
        Loop
        
        ' 3. Make text inside <heading>...</heading> bold and change color to purple
        startPos = InStr(rng.Text, "<heading>")
        Do While startPos > 0
            endPos = InStr(startPos + 9, rng.Text, "</heading>")
            If endPos > startPos Then
                Set headingText = rng.Duplicate
                headingText.Start = rng.Start + startPos + 9 ' After <heading>
                headingText.End = rng.Start + endPos
                
                headingText.Font.Bold = True
                headingText.Font.Color = RGB(128, 0, 128) ' Purple
            End If
            startPos = InStr(endPos + 10, rng.Text, "<heading>")
        Loop
    Next para

    MsgBox "Formatting complete!"
End Sub
