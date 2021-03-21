Sub constSample()
    Static myBlue As Integer = 5
    Range("I11").Select

    Selection.Interior.ColorIndex = myBlue
End SUb