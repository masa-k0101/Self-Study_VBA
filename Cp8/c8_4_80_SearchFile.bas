Sub SearchFile()
    Dim myPath As String
    Dim myFname As String
    Dim i As Integer

    Worksheets("ファイル検索").Activate
    i = 1
    Cells(i, 1).Value = "ファイル名"
    Cells(i, 2).Value = "ファイルサイズ"
    Cells(i, 3).Value = "ファイル作成日付"

    myPath = ActiveWorkbook.Path & "\"

    myFname = Dir(myPath & "*.xls")

    Do While myFname <> ""
        i = i + 1
        Cells(i, 1).Value = myFname
        Cells(i, 2).Value = FileLen(myPath & myFname)
        Cells(i, 3).Value = FIleDateTime(myPath & myFname)

        myFname = Dir()
    Loop
End sub