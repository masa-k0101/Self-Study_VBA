Private Sub Workbook_AddInstall()
    Dim myCB As CommandBar
    Dim myCBCtrl As CommandBarControl, myCBCtrl2 As CommandBarControl

    Set myCB = Application.CommandBars(("Worksheet Menu Bar")

    Set myCBCtrl = myCB.Controls.Add(Type:=msoControlPopup)
    myCBCtrl.Caption = "アドイン"

    Set myCBCtrl2 = myCBCtrl.Controls.Add(Type:=msoControlButton)
    myCBCtrl2.Caption = "マクロ1"
    myCBCtrl2.OnAction = "myMacro1"

    Set myCBCtrl2 = myCBCtrl.Controls.Add(Type:=msoControlButton)
    myCBCtrl2.Caption = "マクロ2"
    myCBCtrl2.OnAction = "myMacro2"

    Set myCBCtrl2 = myCBCtrl.Controls.Add(Type:=msoControlButton)
    myCBCtrl2.Caption = "アドインアンインストール"
    myCBCtrl2.OnAction = "AddUnInstall"
End sub