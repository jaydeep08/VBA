Sub AddSpaceToEndOfContentControls()
    Dim cc As ContentControl
    Dim textLength As Long
    
    ' Loop through all content controls in the document
    For i = ActiveDocument.ContentControls.Count To 1 Step -1
        Set cc = ActiveDocument.ContentControls(i)
        
        ' Store the current length of the text
        textLength = Len(cc.Range.Text)
        
        ' Add a space only if the content control is not empty
        If textLength > 0 Then
            cc.Range.Text = cc.Range.Text & " " ' Add a space
        End If
    Next i

    MsgBox "Spaces have been added to the end of each content control.", vbInformation
End Sub
