Sub ファイル名をまとめて変更する2()
    'アクティブシートのA列に入力されているファイル名を
    'B列に入力されているファイル名に変更する
    'A1セルに、ファイル名を変更したいフォルダのフルパスを入力しておいてください
    'A5セル以下に、現在のファイル名を入力しておいてください
    'B5セル以下に、新しいファイル名を入力しておいてください
    'C5セル以下には実行結果が自動的に入力されます

    Dim fp As String
    Dim i As Long
    Dim fo As String
    Dim fn As String

    'パスを変数に格納
    fp = Range("A1").Value & "\"

    On Error GoTo ERR_HANDL

    Range("C4").Value = "実行結果"

    '5行目から最終行までループ処理を実行
    For i = 5 To Cells(Rows.Count, 1).End(xlUp).Row
        '現在のファイル名を取得
        fo = Cells(i, 1).Value
        '新しいファイル名を取得
        fn = Cells(i, 2).Value
        '新しいファイル名が入力されているときのみ処理を実行
        If fn <> "" Then
            '正常処理の実行結果を先に入力
            Cells(i, 3).Value = _
            "○ファイル名を" & _
            "「" & fo & "」から" & _
            "「" & fn & "」に変更しました。"
            'ファイル名を変更
            Name fp & fo As fp & fn
        End If
    Next i

    Exit Sub
　
    ERR_HANDL:
        Cells(i, 3).Value = _
        "×" & Err.Description & ":" & Err.Number
        Resume Next

End Sub