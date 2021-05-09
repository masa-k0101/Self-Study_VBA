Sub OpenTxtFile3()
    Workbooks.OpenText Filename:="Fuji3.txt", _
    DataType:=xlDelimited, Comma:=True, _
    FieldInfo:=Array(1, 2), Array(2, 1), Array(3, 2), Array(4, 1), _
    Array(5, 1), Array(6, 9), Array(7, 2), Array(8, 2)

End sub