Sub openTxtFile()
    Dim myFName As String

    myFName = Application. _
        GetOpenFilename("テキストファイル　(*.prn; *.txt; *.csv),*prn;*.txt;*.csv")

    If myFName <> "False" Then
        Workbooks.OpenText Filename:=myFName, Comma:=True
    End If
 End SUb