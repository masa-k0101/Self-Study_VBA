Sub MakeLegend()
    With ActiveSheet.ChartObject(1).ChartObject
        .ChartArea.AutoScaleFont = False

        .HasLengend = True

        .Lengend.Position = xlLegendPostitionTop

        .PlotArea_Top = 0
        .PlotArea.Left = 0
        .PlotArea.Width = .ChartArea.Width
        .PlotArea.Height = .ChartArea.Height
    End With
End Sub