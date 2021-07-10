Sub ADOSample1()
    Dim myConnect As New ADODB.Connection
    Dim myRcdSet As New ADODB.Recordset

    myConnect.Open "FIle Name=C:=|Excel2003VBA応用編|test.udl;"

    With myRcdSet
        .ActiveConnection = myConnect
        .SOurce = "Select * From T_申し込み Where コースNo = C001"
        .Open
    End With

    Range("A6").CopyFromRecordset myRcdSet

    myRcdSet.Close
    myConnect.Close
    
    Set myRcdSet = Nothing
    Set myConnect = Nothing
End sub