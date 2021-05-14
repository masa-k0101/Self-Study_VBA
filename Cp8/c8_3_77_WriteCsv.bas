Sub WriteCsv()
    Dim myTxtFile As String
    Dim myFNo As Integer
    Dim myLastRow As String
    Dim i As Integer
    
    Application.ScreenUpdating = False
    
    myTxtFile = ActiveWorkbook.Path & "Numazu.txt"

    Worksheets("郵便番号2").Activate
    myLastRow = Range("A1").CurrentRegion.Rows.Count

    myFNo = FreeFile
    Open myTxtFile For Input As #myFNo

    For i = 1 To myLastRow
        Write #myFNo, Cells(i, 1), Cells(i, 2), Cells(i, 3), Cells(i, 4),_
        Cells(i, 5), Cells(i, 6), Cells(i, 7)
    Next

    Close #myFNo
End sub