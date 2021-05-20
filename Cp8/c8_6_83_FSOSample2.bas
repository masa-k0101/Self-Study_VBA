Sub FSOSample2()
    Dim myFSO As New FileSystemObject

    With myFSO.CreateTextFile("C:|FSOSample1.txt", True)

        .WriteLine "かんたんプログラミングExcel2003VBA応用編" & "作成日：" & Date

        .Close
    End With
End sub