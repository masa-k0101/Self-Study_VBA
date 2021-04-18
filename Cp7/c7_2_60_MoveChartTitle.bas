Sub MoveChartTitle()
    Charts("Graph1").Activate

    With ActiveChart
        .Top = True
        .Left = True
    End With
End Sub