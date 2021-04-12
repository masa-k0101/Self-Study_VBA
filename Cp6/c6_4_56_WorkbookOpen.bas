Private Sub Workbook_Open()
    Dim myCB As CommandBar, myCBCtrl As CommandBarControl

    Set myCB = Application.CommandBars("Worksheet Menu Bar")

    myCB.Controls("書式(&O)").Enabled = False
    myCB.Controls("ツール(&T)").Controls("オプション(&O)...").Enabled = False

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolen)
    Application.CommandBars("Worksheet Menu Bar").Reset
End Sub