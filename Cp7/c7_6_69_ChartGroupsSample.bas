Sub ChartGroupsSample()
    ActiveSheet.ChartObjects(1).Activate

    With ActiveChart.ChartGroups(1)
        .Overlap = 100
        .GapWidth = 50
    End With
End sub