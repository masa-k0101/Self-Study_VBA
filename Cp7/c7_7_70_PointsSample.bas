Sub PointsSample()
    Dim myFormula As Variant
    Dim myRange As Range, i As Long

    ActiveSheet.ChartObject(1).Activate

    With ActiveChart.SeriesCollection(2)

        .Markers = 5

        myFormula = Split(.Formula, ",")

        Set myRange = Range(myFormula(2))

        For i = 1 To myRange.Count
            If myRange.Cells(i).Value = Application.Worksheet.Max(myRange) Then
                Exit For
            End If

        Next
        .Points(i).MarkerSize = 10
    End With
End sub