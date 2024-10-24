Sub DisableContentControllersKeepEditableContent()
    Dim cc As ContentControl

    ' Loop through all content controls in the document
    For Each cc In ActiveDocument.ContentControls
        ' Lock the content control to prevent deletion
        cc.LockContentControl = True
        
        ' Allow editing the content inside the content control
        cc.LockContents = False
    Next cc

    MsgBox "Content controls have been disabled but the inner text is still editable.", vbInformation
End Sub
