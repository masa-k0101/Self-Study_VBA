Sub SeriesCollectionSample()
    Dim mySeries As Series
    Dim myFormula As Variant
    Dim myMsg As String
    
    ActiveSheet.ChartObject(1).Activate

    For Each mySeries In ActiveChart.SeriesCollection

        myFormula = Split(mySeries.Formula, ",")

        myMsg = myMsg & mySeries.Name & " : " & myFormula(2) & vbCrlf
    Next

    MsgBox myMsg
End sub