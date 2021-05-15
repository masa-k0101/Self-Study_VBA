Sub KillFile()
    If Dir("C:DataBook.xls") <> Then
        Kill "C:DataBook.xls"
    Else
        MsgBox "DataBook.xlsは見つかりません"
    End If
End sub