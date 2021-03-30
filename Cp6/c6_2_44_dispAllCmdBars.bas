Sub dispAllComdBars()
    Dim myCB As CommandBar
    Dim i As Integer

    Worksheets("コマンドバー一覧").Activate
    Range("A1").Value = "インデックス番号"
    Range("B1").Value = "名前"
    Range("C1").Value = "種類(数値)"
    Range("D1").Value = "種類(組み込み定数)"

    i = 1
    For Each myCB In Application.CommandBar
        i = i + 1

        Cells(i, 1).Value = myCB.Index

        Cells(i, 2).Value = myCB.Name

        Cells(i, 3).Value = myCB.Type

        Select Case myCB.Type
            Case 0
                Cells(i, 4) = "msoBarTypeNormal"
            Case 1
                Cells(i, 4) = "msoBarTypeMenuBar"
            Case 2
                Cells(i, 4) = "msoBarTypePopup"
        End Select

    Next myCB
 End SUb
