Sub FSOSample6()
    Dim myFSO As New FileSystemObject
    Dim myFld As Folder
    Dim i As Integer

    Worksheets("Sheet2").Activate

    i = 1

    With myFSO.GetFolder("C:|WINNT")
        For Each myFld In .SubFolders
            i = i + 1
            Cells(i, 1).Value = myFld.Name
        Next
    End With
End sub
