Sub trapSample()

    On Error GoTo HandleErr

    ActiveWorkbook.Charts(1).Activate

    ActiveChart.SizeWithWindows = True
    MsgBox "正常に終了しました"
    Exit Sub

HandleErr:
    MsgBox "グラフシートはありません"

End Sub