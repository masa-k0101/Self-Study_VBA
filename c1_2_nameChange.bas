Sub FilenameChange02()
    '指定したフォルダの指定シートのA列をB列のファイル名に変更、エラー制御付き。
 
    Dim File_function As New Scripting.FileSystemObject
    Dim ws01 As Worksheet
    Dim lRow, I As Long
    Dim FolderName, OldFile, NewFile As String
    
    Set ws01 = Worksheets("Sheet2") '保存されている保存先（シート）
    FolderName = "C:\DATA" '保存されている保存先（フォルダー）
    
    lRow = ws01.Cells(Rows.Count, "A").End(xlUp).Row  'A列の最終行を取得
       
    For I = 6 To lRow  'A列の最終行まえ繰り返す
        OldFile = FolderName & "\" & ws01.Cells(I, "A")  'A列から旧ファイル名を取得
        NewFile = FolderName & "\" & ws01.Cells(I, "B")  'B列から新ファイル名を取得
    
        If File_function.FileExists(NewFile) = False Then
               'ファイル名の存在を確認します。既に新ファイル名があれば、変換不可
               
                Name OldFile As NewFile 'ファイル名を変更します。(旧ファイル⇒新ファイル)
                ws01.Cells(I, "C") = "完了"
            Else
                ws01.Cells(I, "C") = "変換不可"
                
        End If
    
    Next I
 
End Sub