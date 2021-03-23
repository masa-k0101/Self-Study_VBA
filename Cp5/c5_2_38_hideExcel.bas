Sub hideExcel()
    Dim myTop As Double, myleft As Double

    Application.WindowState = xlNormal

    myTop = Application.Top
    myleft = Application.Left

    MsgBox "Excelを非表示にする"
    Application.Left = -Application.Width

    MsgBox "Excelを再表示します"
    Application.Top = myTop
    Application.Left = myleft
End SUb