Sub FormatDocumentText()
    Dim doc As Document
    Dim rng As Range
    Dim findColor As Long
    Dim found As Boolean
    
    ' Set the document to the active document
    Set doc = ActiveDocument

    ' Task 1: Make text between <h1> and </h1> bold and blue
    Set rng = doc.Content
    With rng.Find
        .ClearFormatting
        .Text = "<h1>*</h1>"
        .Replacement.ClearFormatting
        .Replacement.Font.Bold = True
        .Replacement.Font.Color = RGB(0, 0, 255) ' Blue color
        .Wrap = wdFindContinue
        .Forward = True
        .Format = True
        .MatchWildcards = True
        
        ' Execute the find and replace
        found = rng.Find.Execute
        Do While found
            ' Select the found range and format it
            rng.Text = Replace(rng.Text, "<h1>", "")
            rng.Text = Replace(rng.Text, "</h1>", "")
            rng.Font.Bold = True
            rng.Font.Color = RGB(0, 0, 255) ' Blue color
            ' Move to the next occurrence
            rng.Start = rng.End
            found = rng.Find.Execute
        Loop
    End With

    ' Task 2: Change specific colored text to black
    Set rng = doc.Content
    findColor = RGB(255, 0, 0) ' Change this to the color you want to replace (e.g., red)
    
    For Each rng In doc.StoryRanges
        Do While rng.Find.Execute(FindText:="", Forward:=True, Wrap:=wdFindContinue, Format:=True)
            If rng.Font.Color = findColor Then
                rng.Font.Color = RGB(0, 0, 0) ' Change to black
            End If
            rng.Start = rng.End ' Move to the next occurrence
        Loop
    Next rng

    ' Task 3: Make text inside <tag> blue
    Set rng = doc.Content
    With rng.Find
        .ClearFormatting
        .Text = "<tag>*</tag>"
        .Replacement.ClearFormatting
        .Replacement.Font.Color = RGB(0, 0, 255) ' Blue color
        .Wrap = wdFindContinue
        .Forward = True
        .Format = True
        .MatchWildcards = True
        
        ' Execute the find and replace
        found = rng.Find.Execute
        Do While found
            ' Select the found range and format it
            rng.Text = Replace(rng.Text, "<tag>", "")
            rng.Text = Replace(rng.Text, "</tag>", "")
            rng.Font.Color = RGB(0, 0, 255) ' Blue color
            ' Move to the next occurrence
            rng.Start = rng.End
            found = rng.Find.Execute
        Loop
    End With

End Sub
