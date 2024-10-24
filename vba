Sub ChangeColorInsideAngularBrackets()
    Dim cc As ContentControl
    Dim startPos As Long, endPos As Long
    Dim txtRange As Range

    ' Loop through all content controls in the document
    For Each cc In ActiveDocument.ContentControls
        If cc.Type = wdContentControlRichText Or cc.Type = wdContentControlText Then
            ' Find positions of < and > within the content control text
            startPos = InStr(cc.Range.Text, "<")
            endPos = InStr(cc.Range.Text, ">")

            ' Ensure both < and > exist and are in the correct order
            If startPos > 0 And endPos > startPos Then
                ' Set the range to the text inside the angular brackets
                Set txtRange = cc.Range.Duplicate
                txtRange.Start = txtRange.Start + startPos
                txtRange.End = txtRange.Start + (endPos - startPos - 1)

                ' Change the font color of the selected range to blue
                txtRange.Font.Color = wdColorBlue
            End If
        End If
    Next cc

    MsgBox "Text inside angular brackets changed to blue!"
End Sub
