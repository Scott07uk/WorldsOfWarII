Attribute VB_Name = "DecrementProtectionTime"
Sub Protectiontime()
    
    main.SQLCon.RecordSource = "SELECT protectionTime FROM tblaccounts WHERE protectionTime > 0"
    main.SQLCon.Refresh
    
    Set rsTime = main.SQLCon.Recordset
    If Not rsTime.EOF Then
        Do While Not rsTime.EOF
            rsTime![Protectiontime] = rsTime![Protectiontime] - 1
            rsTime.Update
            rsTime.MoveNext
        Loop
    End If
    rsTime.Close
End Sub
