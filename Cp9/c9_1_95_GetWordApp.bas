Sub GetWordApp()
    Dim objWord As Word.Application
    Dim myAppOpen As Boolean

    On Error GoTo HandleErr

    Set objWord = GetObject(, "Word.Application")
    myAppOpen = True

MacroContinue:
    If myAppOpen = False Then
        Set objWord = CreateObject("Word.Application")
    End If

   With objWord
        .Visible = True
        .WindowState = wdWindowStateMinimize
       .Documents.Add
    End With

    Set objWord = Nothing

    Exit Sub
    
HandleErr:
    If Err.Number = 429 Then
       myAppOpen = False
       Resume MacroContinue
    End If      
End sub
