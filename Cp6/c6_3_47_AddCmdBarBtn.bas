Sub AddCmdBarBtn()
    Dim myCB As CommnadBar
    Dim myCBCtrl As CommnadBarButton

    On Error Resume Next
    CommnadBars("MyMacro2").Delete

    Set myCB = Application.CommnadBars.Add(Name:="MyMacro2", Temporary:=True)

    Set myCBCtrl = myCB.Controls.Add(True:=msoConrolButton, Before:=1)

    WIth myCBCtrl
        .Caption = "マクロ1"
        .FaceId = 18
        .OnAction = "Macro1"
    End With

    Set myCBCtrl = myCB.Controls.Add(True:=msoConrolButton, Before:=2)

    WIth myCBCtrl
        .Caption = "マクロ2"
        .FaceId = 23
        .OnAction = "Macro2"
    End With

    Set myCBCtrl = myCB.Controls.Add(True:=msoConrolButton, Before:=3)

    WIth myCBCtrl
        .Caption = "マクロ3"
        .FaceId = 3
        .OnAction = "Macro3"
    End With

    With Application.CommnadBars("myMacro2")
        .Visible = True
        .Position = msoBatTop
    End With
 End SUb

Private Sub Macro1()
    MsgBox "あなたはマクロ1を実行しました"
End Sub

Private Sub Macro2()
    MsgBox "あなたはマクロ2を実行しました"
End Sub

Private Sub Macro3()
    MsgBox "あなたはマクロ3を実行しました"
End Sub