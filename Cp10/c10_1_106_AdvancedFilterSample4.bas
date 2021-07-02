Sub AdvancedFilterSample4()
    Dim myRange As Range
    Dim myJoken As Range
    Dim mySheet As Worksheet
    Dim i As Integer

    Range("I5:I15").Clear

    Set myRange = Range("伝票").Offset(, 1).Resize(Range("伝票").Rows.Count, 1)

    myRange.AdvancedFilter xlFilterCopy, , Range("I5"), True

    Set myJoken = Range("I5").CurrentRegion

    For i = 1 To myJoken.Rows.Count - 1

        Set mySheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))

        mySheet.Name = myJoken.Cells(2, 1).Value

        Range("伝票").AdvancedFilter xlFilterCopy, myJoken.Rows("1:2"), mySheet.Range("A2")

        myJoken.Rows(2).Delete xlShiftUp
    Next

    myJoken.Clear
End sub