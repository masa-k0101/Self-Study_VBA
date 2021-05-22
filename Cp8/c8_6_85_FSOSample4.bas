Sub FSOSample4()
    Dim myFSO As New FileSystemObject
    Dim myDrv As Drive
    Dim myMsg As String

    For Each myDrv In myFSO.Drives

        myMsg = myMsg & myDrv.DriveLetter & ": "

        Select Case myDrv.DriveType
            Case 0
                myMsg = myMsg & "不明" & vbCrlf
            Case 1
                myMsg = myMsg & "リムーバブルディスク" & vbCrlf
            Case 2
                myMsg = myMsg & "ハードディスク" & vbCrlf
            Case 3
                myMsg = myMsg & "ネットワークドライブ" & vbCrlf
            Case 4
                myMsg = myMsg & "CD-ROM" & vbCrlf
            Case 5
                myMsg = myMsg & "RAMディスク" & vbCrlf
        End Select
    Next

    MsgBox myMsg 
End sub