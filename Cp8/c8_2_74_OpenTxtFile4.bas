Sub OpenTxtFile4()
    Workbooks.OpenText Filename:="Nyukin.txt", DataType:=xlFixedWidth, _
    FieldInfo:=Array(0, 2), Array(5, 5), Array(13, 1), Array(17, 1), Array(47, 1)
End sub