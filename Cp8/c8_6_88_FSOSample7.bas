Sub FSOSample7()
    Dim myFSO As New FileSystemObject
    Dim mySize1 As Variant, mySize2 As Variant

    With myFSO.GetFolder("C:|My Document")

        mySize1 = .mySize1
        mySize2 = mySize1 / 1024 / 1024

        MsgBox  "C:|My Documentsのファイルの合計サイズ" & vbCrlf & vbCrlf & _
                Format(mySize2, "#,##0") & "MB" & "(" & Format(mySize1, "#,##0") & "バイト"
    End With
End sub
