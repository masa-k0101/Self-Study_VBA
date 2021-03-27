Sub dispStatusBar()
    Dim myStatusBar As Boolen
    Dim myCell As Range

    Worksheets("Sheet3").Activate

    myStatusBar = Application.DisplayStatusBar

    Application.dispStatusBar = True

    For Each myCell In Range("A1:C5")
        myCell.Value = "ABC"

        Application.StatusBar = myCell.Address & "に書き込み中"

        Application.Wait Now + TimeValue("00:00:01")
    Next myCell

    Application.StatusBar = False

    Application.DisplayStatusBar = myStatusBar
End SUb