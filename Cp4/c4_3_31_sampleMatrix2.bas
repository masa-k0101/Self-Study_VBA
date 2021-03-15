Option Explicit
Option Base 1

Sub sampleMatrix()
    Dim myData(3, 2) As String
    
    Dim myVal As Variant

    myData(1, 1) = "大村あつし"
    myData(1, 2) = "フェニックス"

    myData(2, 1) = "井出登志夫"
    myData(2, 2) = "IDE倉庫"

    myData(3, 1) = "増根好夫"
    myData(3, 2) = "大富"

    For Each myVal In myData
        Debug.Print myVal
    Next
End SUb