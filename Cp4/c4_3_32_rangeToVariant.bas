Option Explicit
Option Base 1

Sub rangeToVariant()
    Dim myData As String
    
    Dim r As Integer, c As Integer
    Worksheets("コピー元").Activate

    myData = Range("A1").CurrentRange.Value

    r = UBound(myData, 1)
    c = UBOund(myData, 2)

    Worksheet("コピー元").Activate

    Range(Cells(1, 1), Cells(r, c)).Value = myData
End SUb