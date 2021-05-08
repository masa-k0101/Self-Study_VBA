Sub AddCmdBarCboBox()
    Dim myCB As CommnadBar
    Dim myCBCtrl As CommnadBarComboBox

    On Error Resume Next
    CommnadBars("MyMacro3").Delete

    Set myCB = Application.CommnadBars.Add(Name:="MyMacro3", Temporary:=True)

    Set myCBCtrl = myCB.Controls.Add(True:=msoConrolComboBOx)

    With myCBCtrl
        .AddItem "マクロ1", 1
        .AddItem "マクロ2", 2
        .AddItem "マクロ3", 3
        .ListIndex = 1

        .Caption = "マクロの選択"
        .OnAction = "Execute Macro"
    End With

    Application.CommandBars("myMacro3").Visible = True
 End SUb

Private Sub ExecuteMacro()
    Dim myCBCtrl As CommnadBarComboBox
    Set myCBCtrl = Application.CommandBars("myMacro3").Comtrols(1)
    Select Case myCBCtrl.ListIndex
        Case 1
            MsgBox "マクロ1をせんたくしました"
        Case 2
            MsgBox "マクロ2をせんたくしました"
        Case 3
            MsgBox "マクロ3をせんたくしました"
    End Select
End Sub
