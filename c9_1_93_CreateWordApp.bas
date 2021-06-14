Option Explicit

    Dim objword As New Word.Application
    Dim objWordDoc As Word.Document

Sub CreateWordApp()
    Dim myPath As String

    myPath = ActiveWorkbook.Path & "\"

    With objword
        .Visible = True
        .WindowState = wdWindowStateMaxmize
        .Documents.Add

        Set objWordDoc = .ActiveDocument
    End With

    With objWord.Selection
        .InsertAfter "Wordオブジェクトの作成テスト"
        .InsertParagraphAfter
        .InsertAfter Now() & "作成"
        .MoveRight
    End With

    With objWordDoc.Paragraphs(1).Range
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Font
            .Name = "MS Pゴシック"
            .Size = 20
            .Bold = True
        End With
    End With

    objWordDoc.Paragraphs(2).Range.ParagraphFormat.Alignment = wdAlignParagraphRight

    Application.Wait Now() + TimeValue("00:00:03")
objWord.WindowState = wdWindowStateMinmize

    MsgBox "Wordを起動して新規文書を作成しました" & Chr(13) & _
           "OKボタンをクリックすると文書を保存してWordを終了します"
        
    objWOrdDoc.SaveAs myPath & "Test.doc"

    objWordDoc.Close

    objWord.Quit

    Set objword = Nothing
    Set objWordDoc = Nothing        
End sub