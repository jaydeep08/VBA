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
            ' Loop through each cell in the table where the content control exists
            For Each cellRange In cc.Range.Tables(1).Range.Cells
                Set searchRange = cellRange.Range.Duplicate
                searchRange.End = searchRange.End - 1 ' Exclude end-of-cell marker

                ' Call the helper function to color text inside brackets
                FormatAngularText searchRange
            Next cellRange
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
        With rng.Duplicate
            .Start = rng.Start + startPos
            .End = rng.Start + endPos

            ' Change the font color to blue
            .Font.Color = wdColorBlue
        End With
    End If
End Sub
