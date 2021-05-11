Option Explicit
Option Base 1

Sub ReadTxt()
    Dim myTxtFile As String
    Dim myBuf(7) As String
    Dim i As Integer, j As Integer
    
    Application.ScreenUpdating = False
    
    myTxtFile = ActiveWorkbook.Path & "Fuji.txt"

    Worksheets("郵便番号").Activate

    Open myTxtFile For Input As #1

    Do Until EOF(1)
        Input #1, myBuf(1), myBuf(2), myBuf(3), myBuf(4), myBuf(5), myBuf(6), myBuf(7)

        i = i + 1
        For j = 1 To 7
            Cells(i, j) = myBuf(j)
        Next j
    Loop

    Close #1
End sub