Sub ADOSample1()
    Dim myConnect As New ADODB.Connection
    Dim myRcdSet As New ADODB.Recordset

    myConnect.Open "FIle Name=C:=|Excel2003VBA応用編|test.udl;"

    myRcdSet.Open "T_会員リスト", myConnect

    Range("A6").CopyFromRecordset myRcddSet

    myRcdSet.Close
    myConnect.Close
    
    Set myRcdSet = Nothing
    Set myConnect = Nothing
End sub