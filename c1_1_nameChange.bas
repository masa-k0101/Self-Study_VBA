Sub FilenameChange1()
    '指定したフォルダの指定シートのA列をB列のファイル名に変更します。
 
    Dim ws01 As Worksheet
    Dim lRow, I As Long
    Dim FolderName, OldFile, NewFile As String
    
    Set ws01 = Worksheets("Sheet1") '保存されている保存先（シート）
    FolderName = "C:\DATA" '保存されている保存先（フォルダー）
    lRow = ws01.Cells(Rows.Count, "A").End(xlUp).Row  'A列の最終行を取得
    
    For I = 6 To lRow  '6～最終行まで繰り返す
        OldFile = FolderName & "\" & ws01.Cells(I, "A")
        NewFile = FolderName & "\" & ws01.Cells(I, "B")
        
        MsgBox NewFile
        
        Name OldFile As NewFile 'ファイル名を変更します。
    Next I

End Sub