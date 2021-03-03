Sub trapSample2()

    Dim myRange As Range
    Dim myPropmt As String, myTitle As String

    Worksheets("Sheet3").Activate
    Cells.Clear

    myPropmt = "選択されたセル範囲に「ABC」と表示します" & vbCr & _
     "セル範囲はマウスで選択してください"
    myTitle = "セル範囲入力"

    On Error Resume　Next


    Set myRange = Application.InputBox(Prompt:=myPropmt, Title:=myTitle, Type:=8)

    If myRange Is Nothing Then Exit Sub

    myRange.Value = "ABC"

End Sub
