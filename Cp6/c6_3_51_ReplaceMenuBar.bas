Sub ResetMenuBar()
    Dim myCB As CommandBar

    Set myCB = Application.CommandBars.Add(Name:="User Menu Bar", Position:=Top, MenuBar:=True)

    'コマンドコントロールの作成(サブルーチン)
    S_AddCmdCtrl myCB

    myCB.Visible = True
End Sub