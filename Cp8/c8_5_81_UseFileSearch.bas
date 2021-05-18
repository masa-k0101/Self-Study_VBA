Sub UseFileSearch()
    Dim myFSObj As FileSearch
    Dim i As Integer

    Set myFSObj = Application.FileSearch

    With myFSObj
        .LookIn = ActiveWorkbook.Path
        .Filename = "*.xls"

        If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) > 0 Then

            MsgBox .FoundFiles.Count & " 個のExcelブックが見つかりました"

            For i = 1 To .FoundFiles.Count
                Cells(i, 1).Value = .FoundFiles(i)
            Next i
        Else
            MsgBox "Excelブックはありません"
        End If
    End With
End sub