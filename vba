Sub ChangeColorInsideAngularBrackets()
    Dim cc As ContentControl
    Dim txt As String
    Dim startPos As Long, endPos As Long

    ' Loop through all content controls in the document
    For Each cc In ActiveDocument.ContentControls
        If cc.Type = wdContentControlRichText Or cc.Type = wdContentControlText Then
            txt = cc.Range.Text

            ' Search for the first occurrence of angular brackets <...>
            startPos = InStr(txt, "<")
            endPos = InStr(txt, ">")

            If startPos > 0 And endPos > startPos Then
                ' Apply blue color to the text inside the angular brackets
                With cc.Range.Characters(startPos + 1 To endPos - 1).Font
                    .Color = wdColorBlue ' Change to blue
                End With
            End If
        End If
    Next cc

    MsgBox "Text color inside angular brackets changed to blue!"
End Sub
