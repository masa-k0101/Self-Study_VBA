Sub AddinUnInst()
    MsgBox "アドイン [08章-5]を解除します"
    AddIns("08章-5").Installed = False
End sub