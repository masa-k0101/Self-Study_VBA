Sub ControlNotePad()
    Dim myPath As String
    Dim myID As Double

    myPath = ActiveWorkbook.Path & "|"

    myID = Shell("Notepad.exe", vbNormalFocus)

    SendKeys "%F0", True

    SendKeys myPath & "Report.txt", True
    
    SendKeys "{ENTER}", True
    
    Worksheets("書籍販売").Rnage("販売予測").Copy

    AppActivate myID

    SendKeys "^V", True

    Application.CutCopyMode = False
End sub