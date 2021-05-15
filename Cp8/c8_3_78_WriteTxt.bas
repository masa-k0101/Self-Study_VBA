Sub WriteTxt()
    Dim myTxtFile As String
    Dim myFNo As Integer
    Dim myLastRow As Long
    Dim i As Long
    
    Application.ScreenUpdating = False
    
    myTxtFile = ActiveWorkbook.Path & "Column.txt"

    Worksheets("文書形式2").Activate
    myLastRow = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row

    myFNo = FreeFile
    Open myTxtFile For Output As #myFNo

    For i = 1 To myLastRow
        Print #myFNo, Cells(i, 1)
    Next

    Close #myFNo
End sub