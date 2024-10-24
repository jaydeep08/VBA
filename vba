Sub ChangeColorOnlyInBrackets()
    Dim cc As ContentControl
    Dim cellRange As Range
    Dim angularText As String
    Dim startPos As Long, endPos As Long
    Dim searchRange As Range

    ' Loop through all content controls in the document
    For Each cc In ActiveDocument.ContentControls
        ' Check if the content control is inside a table
        If cc.Range.Tables.Count > 0 Then
            ' Get the range of the content control (which might be in a table cell)
            Set cellRange = cc.Range
            
            ' Exclude the end-of-cell marker from the range
            If cellRange.End > cellRange.Start Then
                cellRange.End = cellRange.End - 1
            End If
            
            ' Call the helper function to color text inside brackets
            FormatAngularText cellRange
        Else
            ' If not inside a table, process the content control directly
            Set searchRange = cc.Range.Duplicate
            FormatAngularText searchRange
        End If
    Next cc

    MsgBox "Text inside angular brackets changed to blue!"
End Sub

' Helper function to find and format text within angular brackets
Sub FormatAngularText(ByVal rng As Range)
    Dim startPos As Long, endPos As Long

    ' Search for the positions of < and >
    startPos = InStr(rng.Text, "<")
    endPos = InStr(rng.Text, ">")

    ' Ensure both < and > are found and are in the correct order
    If startPos > 0 And endPos > startPos Then
        ' Create a range for the text inside the angular brackets
        Dim bracketRange As Range
        Set bracketRange = rng.Duplicate
        bracketRange.Start = rng.Start + startPos
        bracketRange.End = rng.Start + endPos

        ' Change the font color to blue for the text inside the angular brackets
        bracketRange.Font.Color = wdColorBlue
    End If
End Sub
