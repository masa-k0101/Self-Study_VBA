Sub MakeAxesTitle()
    Worksheets("Sheet2").ChartObjects(1).Activate

    With ActiveChart
    .HasTitle = True
    .ChartTitle.Text = "4月度売上高"

        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "アプリケーション"
        End With

        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "個数"
        End With
    End With
End Sub