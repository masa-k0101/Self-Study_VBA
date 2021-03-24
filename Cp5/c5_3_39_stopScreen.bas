Sub stopScreen()
    Dim i As Integer

    Application.ScreenUpdating = False

    For i = 1 To 10
        Worksheets(1).Activate
        Worksheets(2).Activate
    Next Integer

    Application.ScreenUpdating = True
End SUb