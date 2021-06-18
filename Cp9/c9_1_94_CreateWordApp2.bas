Option Explicit

    Dim objword As New Word.Application
    Dim objWordDoc As Word.Document

Sub CreateWordApp2()

    With objword
        .Visible = True
        .WindowState = wdWindowStateMaxmize
        .Documents.Open ActiveWorkbook.Path & "|Report.doc"

        Set objWordDoc = .ActiveDocument

        With .Selection
            .Move Count:=objWordDoc.Characters.Count
            .InsertParagraphAfter
            .InsertAfter "CD販売数"
            .InsertParagraphAfter
            .MoveRight
        End With
    End With

    Worksheets("CD販売").Range("販売枚数").Copy

    With objWord.Selection
        .Paste
        .TypeParagraph
    End With

    Worksheets("CD販売").chartObject(1).Copy

    With objWord
        .Selection.PasteSpecial Placement:=wdInLine, DataType:=wdPasteMetafilePicture
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With
    
    objWord.PrintOut Background:=False

    objWord.Close SaveChange:=False

    objWord.Quit

    Set objword = Nothing
    Set objWordDoc = Nothing        
End sub