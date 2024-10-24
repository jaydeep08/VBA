Sub ChangeColorInBracketsInsideTable()
    Dim cc As ContentControl
    Dim startPos As Long, endPos As Long
    Dim txtRange As Range
    Dim tableCell As Cell

    ' Loop through all content controls in the document
    For Each cc In ActiveDocument.ContentControls
        ' Check if the content control is inside a table
        If Not cc.Range.Tables.Count = 0 Then
            ' Process the content control inside the table
            For Each tableCell In cc.Range.Tables(1).Range.Cells
                ProcessAngularBracketedText tableCell.Range
            Next tableCell
        Else
            ' Process the content control directly if not in a table
            ProcessAngularBracketedText cc.Range
        End If
    Next cc

    MsgBox "Text inside angular brackets changed to blue!"
End Sub

' Helper subroutine to handle angular bracketed text
Sub ProcessAngularBracketedText(ByVal rng As Range)
    Dim startPos As Long, endPos As Long
    Dim txtRange As Range

    ' Find positions of < and > within the range text
    startPos = InStr(rng.Text, "<")
    endPos = InStr(rng.Text, ">")

    ' Ensure both < and > exist and are in the correct order
    If startPos > 0 And endPos > startPos Then
        ' Set the range to the text inside the angular brackets
        Set txtRange = rng.Duplicate
        txtRange.Start = txtRange.Start + startPos
        txtRange.End = txtRange.Start + (endPos - startPos - 1)

        ' Change the font color of the selected range to blue
        txtRange.Font.Color = wdColorBlue
    End If
End Sub
