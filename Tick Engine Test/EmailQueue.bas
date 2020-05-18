Attribute VB_Name = "EmailQueue"
Public Sub clearQueue()
    'This will clear the email queue
    Dim objEmail As Object
    Dim counter As Integer
    
    counter = 0
    
    Set objEmail = CreateObject("Jmail.SMTPMail")
    
    main.SQLCon.RecordSource = "SELECT TOP 5 Emailto, emailsubject, emailbody FROM TblEmailQueue"
    
    main.SQLCon.Refresh
    
    For counter = 1 To 4
        If Not main.SQLCon.Recordset.EOF Then
                With objEmail
                    .serveraddress = "127.0.0.1"
                    .sender = "no-reply@worldsofwar.co.uk"
                    .AddRecipient main.SQLCon.Recordset![emailto]
                    .subject = main.SQLCon.Recordset![emailsubject]
                    .htmlbody = main.SQLCon.Recordset![emailbody]
                    
                    .Execute
                    
                    .ClearRecipients
                End With
                main.SQLCon.Recordset.Delete
                main.SQLCon.Recordset.MoveNext
        End If
    Next
    
    main.SQLCon.Recordset.Close
    
    Set objEmail = Nothing
    
End Sub
