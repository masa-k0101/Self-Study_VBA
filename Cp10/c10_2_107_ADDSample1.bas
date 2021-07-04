Sub ADOSample1()
    Dim myConnect As New ADODB.Connection
    Dim myRcdSet As New ADODB.Recordset
    Dim myProvider As String, mySource As String

    myProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"
    mySource = "Data Source=C:=|Excel2003VBA応用編|会員管理.mdb;"

    myConnect.Open myProvider & mySource

    myRcdSet.Open "T_会員リスト", myConnect

    Range("A6").CopyFromRecordset myRcddSet

    myRcdSet.Close
    myConnect.Close
    
    Set myRcdSet = Nothing
    Set myConnect = Nothing
End sub