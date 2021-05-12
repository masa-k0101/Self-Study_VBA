Sub ReadTxt2()
    Dim myTxtFile As String
    Dim myFNo As Integer
    Dim myBuf As String
    Dim i As Integer
    
    Application.ScreenUpdating = False
    
    myTxtFile = ActiveWorkbook.Path & "\WordPro.txt"

    Worksheets("文書形式").Activate

    myFNo = FreeFile
    Open myTxtFile For Input As #myFNo

    Do Until EOF(myFNo)
        Line Input #myFNo, myBuf

        i = i + 1
            Cells(i, j) = myBuf
    Loop

    Close #myFNo
End sub