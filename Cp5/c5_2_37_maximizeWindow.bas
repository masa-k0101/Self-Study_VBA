Sub maximizeWindow()
    With ActiveWindow
        .WindowState = xlNormal
        .Top = 0
        .Left = 0
        .Height = Application.UsableHeight
        .Width = Application.UsableWidth
    End With
End SUb