Sub MakeChart()
    Dim mySource As Range

    Set mySource = Range("B2").CurrentRegion

    Chart.Add

    ActiveChart.SetSourceData Source:=mySource, PlotBy:=xlColumns
    
End Sub