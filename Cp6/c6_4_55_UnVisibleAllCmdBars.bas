Sub UnVisibleAllCmdBar()
    Dim myCB As CommandBar

    On Error Resume Next

    For Each myCB In Application.CommandBars
        myCB.Visible = False
    Next myCB

    On Error Goto 0

    Application.CommandBars("Worksheet Menu Bar").Enabled = False

End Sub