Sub openFIle()

    Dim myFD As Variant, myFN As Variant
    Dim myPrompt As String, myMsg As String
    Dim myBuf As String

    MsgBox "FDのルートディレクトリにtxtファイルを準備してください"
        & vbCr & "ファイル名は任意で構いません"

InputFD:
    myPrompt = "フロッピーディスク名を入力してください"
    myFD = Application.InputBox(Prompt:=myPrompt, Defalut:="A")
    If varType(myFD) <> vbString THen Exit Sub

InputFN:
    myPrompt = "ファイル名を入力してください"
    myFN = Application.InputBox(Prompt:=myPrompt)
    If varType(myFN) <> vbString THen Exit Sub

    On Error  GoTo HandleErr

    Open myFD & ":\" & myFN For Input As #1
    
    Do Until EOF(1)
        Line Input #1, myBuf
    Loop

    MsgBox "正常に終了しました"
    Close #1

    Exit Sub

HandleErr:
    Select Case Err.Number
        Case 53
            MsgBox Err.Description & vbCr & "ファイル名を再入力してください"
            Resume InputFN
        Case 55
            MsgBox Err.Description
            Resume Next
        Case 68, 75, 76
            MsgBox Err.Description & vbCr & vbCr & "無効のドライブを指定しました" & _
            "フロッピードライブを再入力してください"
            Resume InputFD
        Case 52, 71
            MsgBox Err.Description & vbCr & "フロッピーをセットして処理を続行しますか"
            If MsgBox(myMsg, vbExclamation + vbYesNo) = vbYes Then 
                Resume
            Else
                Exit Sub
            End If
    End Select
End Sub