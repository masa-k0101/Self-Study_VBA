Sub displayErr()

    Dim myMsg As String

    On Error  GoTo HandleErr

    Range("B3").Value = Range("B1").Value / Range("B2").Value
    
    Exit Sub

HandleErr:
    myMsg = "エラー番号：" & Err.Number & vbCrLf &  "エラー内容：" & Err.Description

    MsgBox myMsg

    Range("B3").Value = 0

End Sub