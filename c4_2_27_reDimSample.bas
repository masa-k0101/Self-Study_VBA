Option Explicit
Option Base 1

Sub reDimSample()
    Dim myName() As String
    
    Dim r As Long, i As Long

    Worksheet("3年A組").Activate

    'データが入力された最終行を算出
    r = Range("A65536").End(xlUp).Row

    ReDim myName(r)

    For i = 1 To r
        myName(i) = Cells(i, 1).Value
    Next

    For i = 1 To r
        Debug.Print myName(i)
    Next
End SUb