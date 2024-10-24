Sub ChangeColorInsideAngularBrackets()
    Dim cc As ContentControl
    Dim txt As String
    Dim startPos As Long, endPos As Long

    ' Loop through all content controls in the document
    For Each cc In ActiveDocument.ContentControls
        If cc.Type = wdContentControlRichText Or cc.Type = wdContentControlText Then
            txt = cc.Range.Text

            ' Find the position of the first < and >
            startPos = InStr(txt, "<")
            endPos = InStr(txt, ">")

            ' Ensure both < and > are found, and they are in correct order
            If startPos > 0 And endPos > startPos Then
                ' Select the range inside the angular brackets
                With cc.Range.Characters(startPos + 1 To endPos - 1).Font
                    .Color = wdColorBlue ' Set text color to blue
                End With
            End If
        End If
    Next cc

    MsgBox "Text color inside angular brackets changed to blue!"
End Sub
