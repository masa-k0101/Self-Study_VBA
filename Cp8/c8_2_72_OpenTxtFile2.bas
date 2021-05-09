Sub OpenTxtFile2()
    Workbooks.OpenText Filename:="Fuji2.txt", DataType:=xlDelimited,_
    ConsecutiveDelimiter:=True, Comma:=True, Space:=True

End sub