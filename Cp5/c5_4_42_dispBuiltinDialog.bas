Sub dispBuiltinDialog()
    Dim myRtn As Boolen

    myRtn = Application.Dialog(xlDialogOptionView).Show

    If myRtn = False Then
        MsgBox "「キャンセル」が選択されました" & vbCrlf "処理を終了します"
        Exit Sub
    End If

    MsgBox "処理を続行します"

    '処理を続行します
    
 End SUb