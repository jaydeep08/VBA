Sub RemoveContentControllersKeepText()
    Dim cc As ContentControl

    ' Loop through the content controls in reverse to safely remove them
    For i = ActiveDocument.ContentControls.Count To 1 Step -1
        Set cc = ActiveDocument.ContentControls(i)

        ' Replace the content control with its text
        cc.Range.Select
        cc.Range.Text = cc.Range.Text

        ' Delete the content control
        cc.Delete
    Next i

    MsgBox "All content controls have been removed, but the text remains.", vbInformation
End Sub
