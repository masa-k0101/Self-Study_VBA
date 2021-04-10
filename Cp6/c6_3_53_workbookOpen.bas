Private Sub workbook_Open()
    Dim myCBCtrl As CommandBarButtom

    Set myCBCtrl = Application.CommandBars("Cell").Controls.Add_
        (Type:=msoControlButtom, Id:=872, Before:=8, Temporary:=True)
    myCBCtrl.Caption = "書式のクリア(&F)"
End Sub

Private Sub workbook_BeforeClose(Cancel As Boolen)
    Application.CommandBars("Cell").Reset
End Sub