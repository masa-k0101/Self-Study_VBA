Sub FSOSample5()
    Dim myFSO As New FileSystemObject

    If myFSO.Drive("A").IsReady = True Then
        FileCopy "C:|Excel2003VBA応用編|FUji.txt", "A:|Fuji.txt"
    Else
        MsgBox "フロッピーが挿入されていません"
    End If
End sub