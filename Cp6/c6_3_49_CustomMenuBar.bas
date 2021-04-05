Sub CustomMenuBar()
    Dim myCB As CommandBar, myCBCtrl As CommandBarControl

    Set myCB = Application.CommandBars("Worksheet Menu Bar")

    For Each myCBCtrl In myCB.Controls
        myCBCtrl.Delete
    Next myCBCtrl

    Set myCBCtrl = myCB.Controls.Add(Type:=msoControlPopup)
    myCBCtrl.Caption = "ファイル(&F)"

    Set myCBCtrl = myCB.Controls("ファイル(&F)").Controls.Add(Type:=msoControlButton)
    myCBCtrl.Caption = "保存(&S)"
    myCBCtrl.OnAction = "S_SaveBook"


End Sub

Private Sub S_SaveBook()
    'ActiveWorkbook.Save
    MsgBox "「ファイル」->「保存」コマンドが選択されました"
End Sub