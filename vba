Sub AppendNewLineToContentControls()
    Dim cc As ContentControl

    ' Loop through all content controls in the document
    For Each cc In ActiveDocument.ContentControls
        ' Check if the content control contains text
        If cc.Range.Text <> "" Then
            ' Move to the end of the content control's range and insert a new line
            With cc.Range
                .Collapse Direction:=wdCollapseEnd  ' Move to the end of the content control
                .InsertAfter vbCrLf  ' Append a new line
            End With
        End If
    Next cc

    MsgBox "New lines appended to all content controls.", vbInformation
End Sub
