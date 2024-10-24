Sub RemoveContentControlsByAddingNewLine()
    Dim cc As ContentControl

    ' Loop through all content controls in the document
    For Each cc In ActiveDocument.ContentControls
        ' Add a newline at the end of the content control's text
        cc.Range.Text = cc.Range.Text & vbCrLf
    Next cc

    MsgBox "Content controls removed by adding new lines.", vbInformation
End Sub
