Sub AdvancedFilterSample2()

    Range("伝票").AdvancedFilter　xlFilterCopy, Range("条件範囲"), _
    Worksheets("抽出範囲").Range("A2")

End sub