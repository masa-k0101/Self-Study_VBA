Sub stopAlertMsg()
    Application.DisplayAlerts = False

    Worksheets.Add.Name = "Dummy"
    MsgBox "ワークシート「Dummy」を削除しました" & vbCrLf & "これから「Dummy」を削除します"

    Worksheets("Dummy").Delete

    Application.DisplayAlerts = True
End SUb