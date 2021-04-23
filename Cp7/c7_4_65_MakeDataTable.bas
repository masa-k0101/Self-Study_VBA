Sub MakeDataTable
    With ActiveSheet.ChartObject(1).Chart
        .HasLegend = False

        .HasDataTable = True

        With .DataTable
            .ShowLegendKey = True
            .Font.Bold = True
        End With
    End With
End sub