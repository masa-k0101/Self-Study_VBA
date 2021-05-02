Sub MakeLegend()
    With ActiveSheet.ChartObject(1).ChartObject
        .ChartArea.AutoScaleFont = False

        .HasLegend = True

        .Lengend.Position = xlLegendPositionTop

        .PlotArea_Top = 0
        .PlotArea.Left = 0
        .PlotArea.Width = .ChartArea.Width
        .PlotArea.Height = .ChartArea.Height
    End With
End Sub
