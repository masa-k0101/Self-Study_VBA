Sub AdvancedFilterSample3()
    Dim myRange As Range

    Range("I5:I15").Clear

    Set myRange = Range("伝票").Offset(, 1).Resize(Range("伝票").Rows.Count, 1)

    myRange.AdvancedFilter xlFilterCopy, , Range("I15"), True
End sub