Sub AutoFilterModeSample()
    With ActiveSheet
        MsgBox "フィルタモードは" & .AutoFilterMode & "です"
        .AutoFilterMode = Not .AutoFilterMode
    End With
End sub