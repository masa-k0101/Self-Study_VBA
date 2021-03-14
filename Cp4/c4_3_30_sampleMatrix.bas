Option Explicit
Option Base 1

Sub sampleMatrix()
    Dim myData(3, 2) As String
    
    Dim i As Integer, j As Integer

    myData(1, 1) = "大村あつし"
    myData(1, 2) = "フェニックス"

    myData(2, 1) = "井出登志夫"
    myData(2, 2) = "IDE倉庫"

    myData(3, 1) = "増根好夫"
    myData(3, 2) = "大富"

    For i = 1 To 3
        For j = 1 To 2
            Debug.Print myData(i, j)
        Next j
    Next i
End SUb