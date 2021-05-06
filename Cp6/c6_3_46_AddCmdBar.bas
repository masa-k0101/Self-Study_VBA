Sub AddCmdBar()
    Dim myCB As CommandBar

    On Error Resume Next
    CommandBars("MyMacro").Delete

    Set myCB = Application.CommandBars.Add(Name:="MyMacro")

    With myCB
        .Controls.Add Type:=msoControlButton, Id:=2520, Before:=1
        .Controls.Add Type:=msoControlButton, Id:=23, Before:=2
        .Controls.Add Type:=msoControlButton, Id:=3, Before:=3
        .Visible = True
    End With
 End SUb
