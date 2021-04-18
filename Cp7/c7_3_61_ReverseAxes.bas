Sub ReverseAxes()
    Worksheets("Sheets").ChartObjects(1).Activate
    ActiveChart.Axes(xlCategory).ReversePlotOrder = True

End Sub