Sub FilenameChange04()  '指定した新ファイル名を変換します。
 
    Dim File_function As New Scripting.FileSystemObject
    Dim RC As Integer
    Dim lRow, I As Long
    Dim FolderName, OldFile, NewFile As String
    Dim ws01 As Worksheet
    
    Set ws01 = Worksheets("Sheet3")
    
    lRow = ws01.Cells(Rows.Count, "A").End(xlUp).Row  'A列の最終行を取得
    
    For I = 6 To lRow
        If IsEmpty(ws01.Range("B" & I)) = True Then
            MsgBox "新ファイル名を指定していないセルがあります。"
            Exit Sub
        End If
    Next I
    
    
    RC = MsgBox("選択したファイル名を変更しますか？", vbYesNo + vbQuestion, "確認")
              'ファイル名変換を実行するか確認します。
    If RC = vbNo Then
                MsgBox ("ファイル名変換をキャンセルしました。")
                Exit Sub   'プログラムを中断
    End If
    
    FolderName = ws01.Range("A3") '保存されている保存先（フォルダーパス）
       
    For I = 6 To lRow
    
        OldFile = FolderName & "\" & ws01.Cells(I, "A") 'A列から旧ファイル名を取得
        NewFile = FolderName & "\" & ws01.Cells(I, "B") 'B列から新ファイル名を取得
    
        If File_function.FileExists(NewFile) = False Then
               
                Name OldFile As NewFile 'ファイル名を変更します。(旧ファイル⇒新ファイル)
                ws01.Cells(I, "C") = "完了"
            Else
                ws01.Cells(I, "C") = "変換不可"
                
        End If
    
    Next I
 
End Sub