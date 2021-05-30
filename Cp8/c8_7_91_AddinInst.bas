Sub AddinInst()
    If AddIns("08章-5").Installed = False Then
        MsgBox "アドイン [08章-5]を組み込みます"
        AddIns("08章-5").Installed = True
    Else
        MsgBox "アドイン [08章-5]を組み込みまれています"
    End If
End sub