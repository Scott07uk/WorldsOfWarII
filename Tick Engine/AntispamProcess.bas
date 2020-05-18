Attribute VB_Name = "AntispamProcess"
Sub processspam()
    
    main.SQLCon.RecordSource = "SELECT tblaccounts.ticksmsg FROM tblaccounts"
    main.SQLCon.Refresh
    
    Set rsMsg = main.SQLCon.Recordset
    
    If Not rsMsg.EOF Then
        Do While Not rsMsg.EOF
            rsMsg![ticksmsg] = 0
            rsMsg.Update
            rsMsg.MoveNext
        Loop
    End If
    
    rsMsg.Close
    Set rsMsg = Nothing

End Sub
