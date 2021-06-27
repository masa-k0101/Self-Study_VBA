Sub AutoFilterSample3()
    Dim myCell As Range
    Dim myCode As Variant

    myCode = Application.InputBox("末尾の数字を入力してください")

    If myCode = False Then Exit Sub

    For Each myCell In Range("伝票").Offset(1).Resize(Range("伝票").Rows.Count - 1, 1)
        myCell.Value = "|" & myCell.Value
    Next

    Selection.AutoFilter Field:=1, Criteria1:="=*" & myCode
End sub