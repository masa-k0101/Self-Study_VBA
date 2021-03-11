Option Explicit

Sub oneWeek()
    Dim myWeek(6) As String
    Dim i As Integer

    myWeek(0) = "日曜日"
    myWeek(1) = "月曜日"
    myWeek(2) = "火曜日"
    myWeek(3) = "水曜日"
    myWeek(4) = "木曜日"
    myWeek(5) = "金曜日"
    myWeek(6) = "土曜日"

    For i = 0 To 6
        Debug.Print myWeek(i)
    Next
End SUb
