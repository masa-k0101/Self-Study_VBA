Private Sub workbook_BeforeRightClick(ByVal Target As Excel.Range, Cancel As Boolen)
    Dim myCB As CommandBar
    Dim myCBCtrl As CommandBarControl

    Set myCB = Application.CommandBars.Add(Nmae:="User Short Menu", Position:=msoBarPopup)

    'コマンドコントロールの作成(サブルーチン)
    S_AddCmdCtrl myCB

    myCB.ShowPopup

    CommandBars("User Short Menu").Delete

    Cancel = True

End Sub