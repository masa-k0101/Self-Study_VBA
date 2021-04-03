Sub displayCommandId()
    Dim myBar As CommandBar, myBarName As String
    DIm myCtrl As CommandBarControl
    Dim myCtrl1 As CommandBarControl, myCtrl2 As CommandBarControl

    Range("A1").Select
    ActiveCell.Value = "Index"
    ActiveCell.Offset(, 1) = "Name"

    ActiveCell.Offset(, 2) = "Id"
    ActiveCell.Offset(, 3) = "Caption"

    ActiveCell.Offset(, 4) = "Id"
    ActiveCell.Offset(, 5) = "caption"

    ActiveCell.Offset(, 6) = "Id"
    ActiveCell.Offset(, 7) = "caption"

    ActiveCell.Offset(1).Select

    For Each myBar In Application.CommandBars
        ActiveCell.Value = myBar.Index
        ActiveCell.Offset(, 1) = myBar.Name

        myBarName = myBar.Name
        For Each myCtrl In Application.CommandBars(myBarName).Controls
            ActiveCell.Offset(, 2) = myCtrl.Id
            ActiveCell.Offset(, 3) = myCtrl.Caption

            On Error Resume Next

            With Application.CommandBars(myBarName).Controls(myCtrl.Caption)
                For Each myCtrl1 In Controls
                    ActiveCell.Offset(, 4) = myCtrl1.Id
                    ActiveCell.Offset(, 5) = myCtrl1.Caption

                With Application.CommandBars(myBarName).Controls(myCtrl.Caption)_
                .Controls(myCtrl1.Caption)
                    For Each myCtrl2 In Controls
                        ActiveCell.Offset(, 6) = myCtrl2.Id
                        ActiveCell.Offset(, 7) = myCtrl2.Caption
                        ActiveCell.Offset(1).Select
                    Next myCtrl2
                End With

                Next myCtrl1
            End With

            ON Error Goto 0

        Next myCtrl
    Next myBar
 End SUb