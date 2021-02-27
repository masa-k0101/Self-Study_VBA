Sub FilenameChange03()  'ダイアログボックスを表示してファイル名一覧を作成します。
 
    Dim File_function As New Scripting.FileSystemObject
    Dim lRow, I, F As Long
    Dim FolderName, OldFile, NewFile As String
    Dim FileName As Variant
    Dim ws01 As Worksheet
    
    Set ws01 = Worksheets("Sheet3")
    
    FileName = Application.GetOpenFilename(MultiSelect:=True)
            'ダイアログボックスが表示（MultiSelect:=Trueでファイルを複数選択）
    
    If FileName(1) <> False Then
         
                FolderName = File_function.GetParentFolderName(FileName(1))
                '選択した最初のファイル名からフォルダーまでのルートを取得する
            
            Else
                MsgBox "作業をキャンセルされました"
                Exit Sub 'プログラムを終了
                
    End If
    
    lRow = ws01.Cells(Rows.Count, "A").End(xlUp).Row  'A列の最終行を取得
       
    ws01.Range("A6:A" & lRow + 1).ClearContents 'A列のデータ（文字列のみ）をクリアー
                   
    F = 1  '選択ファイルの１件目を設定
       
    For I = 6 To 5 + UBound(FileName) '選択したファイルの数を繰り返す。（最大値）
    
        ws01.Range("A" & I) = File_function.GetFileName(FileName(F))
        'ファイル名を順番にA列（セル）へ転記します。
        
        F = F + 1
        '次のファイル名を指定するために＋１加算する。
    Next I
 
    ws01.Range("A3") = FolderName  '選択したフォルダーバスをセル「A3]へ転記
 
End Sub