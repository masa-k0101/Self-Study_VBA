Sub FSOSample3()
    Dim myFSO As New FileSystemObject
    Dim myDS1 As Variant
    Dim myDS2 As Variant

    With myFSO.GetDrive("A")

        myDS1 = .TotalSize

        myDS2 = .AvailableSpace

    End With

    MsgBox  "使用領域：" & Format(myDS1 - myDS2, "#,##0") & vbCrlf & _
            "空き領域：" & Format(myDS2, "#,##0") & vbCrlf & vbCrlf & _
            "総容量：" & Format(myDS1, "#,##0") & vbCrlf 
End sub