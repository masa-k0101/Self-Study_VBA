Sub MakeDataLabels()
    Dim myRange As Range
    Dim i As Long

    Set myRange = Range("A2", Range("A2").End(xlDown))

    ActiveSheet.ChartObject(1).Activate

    ActiveChart.ApplyDataLabels

    For i To myRange.Count
        ActiveChart.SeriesCollection(1).Points(i).DataLabel.Text = myRange.Cells(i).Value
    Next i
End sub