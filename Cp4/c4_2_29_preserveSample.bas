Option Explicit
Option Base 1

Sub preserveSample()
    Dim myName() As String
    
    Dim r As Long, r2 As Long
    Dim i As Long

    Worksheet("3年A組").Activate

    r = Range("A65536").End(xlUp).Row

    ReDim myName(r)

    For i = 1 To r
        myName(i) = Cells(i, 1).Value
    Next

    Worksheet("3年B組").Activate

    r2 = Range("A65536").End(xlUp).Row

    ReDim preserve myName(r + r2)

    For i = r + 1 To r + r2
        myName(i) = Cells(i - r, 1).Value
    Next

    For i = LBound(myName) To UBound(myName)
        Debug.Print myName(i)
    Next
    
End SUb
