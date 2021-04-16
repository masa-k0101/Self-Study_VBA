Sub MakeChartTitle()
    Charts("Graph1").Activate

    With ActiveChart
        .HasTitle = True
        .ChartTitle = True
    End With
End Sub