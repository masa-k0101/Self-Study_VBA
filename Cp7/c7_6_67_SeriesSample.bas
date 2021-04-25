Sub SeriesSample()
    ActiveSheet.ChartObject(1).Activate

    With ActiveChart.SeriesCollection(3)
        .ChartTyep = xlLineMarkers
        .AxisGroup = xlSecondary
    End With
End sub