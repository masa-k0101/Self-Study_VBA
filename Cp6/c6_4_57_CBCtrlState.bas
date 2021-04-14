Private Sub CBCtrl_State()
    Dim myCBCtrl As CommandBarButton

    Set myCBCtrl = CommandBars("MyMacro4").COntrols("CheckMark").Controls("CheckMarkOff")

    If myCBCtrl.State = msoButtonDown Then
        myCBCtrl.State = msoButtonUp
        MsgBox "チェックマークをオフしました"
    Else
        myCBCtrl.State = msoButtonDown
        MsgBox "チェックマークをオンしました"
    End If
End Sub