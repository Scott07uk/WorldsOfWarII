Attribute VB_Name = "DelInactives"
Public Sub emailInactives()
Dim intNextEmail As Integer
Print #1, "+------------------------------------------------------+" & vbCrLf
Print #1, "| Emailing accounts that have been inactive for 7 days |" & vbCrLf
Print #1, "+------------------------------------------------------+" & vbCrLf & vbCrLf

main.SQLCon.RecordSource = "SELECT email, username, lastlogin FROM tblaccounts WHERE (lastlogin = 168) OR (lastlogin = 336) AND holiday = 0"

main.SQLCon.Refresh

main.SQLCon2.RecordSource = "SELECT tblemailqueue.* FROM tblemailqueue order by emailID DESC;"

main.SQLCon2.Refresh

If main.SQLCon2.Recordset.EOF Then
        intNextEmail = 1
    Else
        intNextEmail = CInt(main.SQLCon2.Recordset![EmailID]) + CInt(1)
End If

Do While Not main.SQLCon.Recordset.EOF
    Print #1, main.SQLCon.Recordset![UserName] & " has been inactive for 7 days, email has been added to the queue" & vbCrLf
    main.SQLCon2.Recordset.AddNew
    main.SQLCon2.Recordset![EmailID] = intNextEmail
    main.SQLCon2.Recordset![emailto] = main.SQLCon.Recordset![email]
    main.SQLCon2.Recordset![emailsubject] = "Your Worlds of War Account"
    'message to be spell checked
    main.SQLCon2.Recordset![emailbody] = "<font face=arial size=2>" & main.SQLCon.Recordset![UserName] & "<BR>Your account with Worlds of War II has been inactive for " & (CLng(main.SQLCon.Recordset![lastlogin]) / 24) & " days now, If you do not wish you account to be deleted then please login to <A HREF=http://www.worldsofwar.co.uk target=_blank>www.worldsofwar.co.uk</A> and your account will be safe.  If you no longer wish to play Worlds of War II, you need not do anything, your account will be queued for deletion in the next few days.  <BR /><BR /><I>Worlds of War Creators</I> </font>"
    'end message
    
    'save the email in the queue
    main.SQLCon2.Recordset.Update
    
    main.SQLCon.Recordset.MoveNext
Loop

Print #1, "All accounts that have been inactive for 7 days have been emailed.  " & vbCrLf & vbCrLf & vbCrLf

main.SQLCon2.Recordset.Close
main.SQLCon.Recordset.Close

End Sub


Public Sub incLastLogin()
On Error Resume Next
Print #1, "+-----------------------------------------+" & vbCrLf
Print #1, "| Incrementing last login on all accounts |" & vbCrLf
Print #1, "+-----------------------------------------+" & vbCrLf & vbCrLf

'method to change the last login time
    main.SQLCon.RecordSource = "SELECT tblaccounts.lastlogin, ticksmsg, username, protectiontime, email, dob, continent, countryid, printed FROM tblaccounts;"
    
    main.SQLCon.Refresh
    
    Do While Not main.SQLCon.Recordset.EOF
        Call main.SQLCon.Recordset.Update("lastlogin", CLng(CLng(main.SQLCon.Recordset![lastlogin]) + 1))
        Call main.SQLCon.Recordset.Update("ticksmsg", 0)
        If main.SQLCon.Recordset![ProtectionTime] > 0 Then
                Call main.SQLCon.Recordset.Update("protectiontime", main.SQLCon.Recordset![ProtectionTime] - 1)
        End If
        Print #1, main.SQLCon.Recordset![UserName] & " incremented successfully to " & CLng(main.SQLCon.Recordset![lastlogin])
        
        If CBool(main.SQLCon.Recordset![printed]) = False Then
                'Printer.Print ("USER SIGNUP FORM")
                'Printer.Print ("USERNAME: " & main.SQLCon.Recordset![UserName])
                'Printer.Print ("   EMAIL: " & main.SQLCon.Recordset![email])
                'Printer.Print ("     DOB: " & main.SQLCon.Recordset![dob])
                'Printer.Print ("LOCATION: " & main.SQLCon.Recordset![continent] & ":" & main.SQLCon.Recordset![countryid])
                'Printer.EndDoc
                Call main.SQLCon.Recordset.Update("printed", 1)
        End If
        main.SQLCon.Recordset.MoveNext
    Loop
    
    main.SQLCon.Recordset.Close
        
    Print #1, "All accounts have been incremented" & vbCrLf & vbCrLf

End Sub

Public Sub delInactiveAccounts()
Dim lngCounter As Long
Print #1, "+-----------------------------+" & vbCrLf
Print #1, "| Deleteing inactive accounts |" & vbCrLf
Print #1, "+-----------------------------+" & vbCrLf & vbCrLf

main.SQLCon.RecordSource = "SELECT tblaccounts.* FROM tblaccounts WHERE lastlogin > 480;"

main.SQLCon.Refresh

Do While Not main.SQLCon.Recordset.EOF
    Printer.FontName = "CP 860 7.6 CPI"
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.Print "WoWII Account Deletion"
    Printer.FontBold = False
    Printer.FontUnderline = False
    Printer.Print
    Printer.FontName = "CP 866 7.6 CPI"
    Printer.FontUnderline = True
    Printer.Print "User Details:"
    Printer.FontUnderline = False
    Printer.Print
    Printer.FontName = "CP 866 19 CPI"
    Printer.Print "Username: " & main.SQLCon.Recordset![UserName]
    Printer.Print "E-mail: " & main.SQLCon.Recordset![email]
    Printer.Print "County: " & main.SQLCon.Recordset![county]
    Printer.Print "Country: " & main.SQLCon.Recordset![country]
    Printer.Print "DOB: " & main.SQLCon.Recordset![dob]
    Printer.Print "Location: " & main.SQLCon.Recordset![continent] & ":" & main.SQLCon.Recordset![countryid]
    Printer.Print "Suspended: " & CBool(main.SQLCon.Recordset![suspended])
    Printer.Print "No Login: " & CBool(main.SQLCon.Recordset![firstlogin])
    Printer.Print "Score: " & main.SQLCon.Recordset![score]
    Printer.Print
    Printer.FontName = "CP 866 7.6 CPI"
    Printer.FontUnderline = True
    Printer.Print "Other Details:"
    Printer.FontUnderline = False
    Printer.Print
    Printer.FontName = "CP 866 19 CPI"
    Print #1, main.SQLCon.Recordset![UserName] & " is inactive and is being deleted"
    'remove the vote for the power base
    
    'check to see if the person is voting for themselfs
    If main.SQLCon.Recordset![UserName] = main.SQLCon.Recordset![Vote] Then
            'the person is voting for themselfs so do nothing
        Else
            'open the other persons account and decrement their votes
            main.SQLCon2.RecordSource = "SELECT tblaccounts.numvotes FROM tblaccounts WHERE username = '" & main.SQLCon.Recordset![Vote] & "';"
            
            main.SQLCon2.Refresh
            
            If main.SQLCon2.Recordset.EOF Then
                    'there has been an error just delete the account
                Else
                    Call main.SQLCon2.Recordset.Update("numvotes", CInt(main.SQLCon2.Recordset![NumVotes]) - 1)
                    'vote has been decremembted
            End If
            main.SQLCon2.Recordset.Close
    End If
    Print #1, "Vote for power base removed"
    
    'clear out all the messages to this person
    main.SQLCon2.RecordSource = "SELECT * FROM tblmessages WHERE tocontinent = " & main.SQLCon.Recordset![continent] & " AND tocountry = " & main.SQLCon.Recordset![countryid]
    
    main.SQLCon2.Refresh
    lngCounter = 0
    Do While Not main.SQLCon2.Recordset.EOF
        'delete the message
        main.SQLCon2.Recordset.Delete
        'audit the delete
        Print #1, "Message to " & main.SQLCon.Recordset![UserName] & " deleted." & vbCrLf
        'mode to next message
        main.SQLCon2.Recordset.MoveNext
        lngCounter = lngCounter + 1
    Loop
    'clean up
    main.SQLCon2.Recordset.Close
    'messages to deleted
    Printer.Print "Msgs Deleted: " & lngCounter
    Print #1, "All messages to " & main.SQLCon.Recordset![UserName] & " deleted." & vbcrfl & vbCrLf
    Print #1, "Deleting messages from " & main.SQLCon.Recordset![UserName] & vbCrLf
    
    'clear out the messages from
    main.SQLCon2.RecordSource = "SELECT tblmessages.* FROM tblmessages WHERE fromcontinent = " & main.SQLCon.Recordset![continent] & " AND fromcountry = " & main.SQLCon.Recordset![countryid]
    
    main.SQLCon2.Refresh
    
    Do While Not main.SQLCon2.Recordset.EOF
        'delete message
        main.SQLCon2.Recordset.Delete
        'audit the delete
        Print #1, "Message from " & main.SQLCon.Recordset![UserName] & " deleted." & vbCrLf
        'move to next one
        main.SQLCon2.Recordset.MoveNext
    Loop
    
    'clear up
    main.SQLCon2.Recordset.Close
    Print #1, "All messages from " & main.SQLCon.Recordset![UserName] & " deleted." & vbCrLf & vbCrLf
    
    
    'clear out all the news for this person
    Print #1, "Clearing out all news items.  " & vbCrLf
    main.SQLCon2.RecordSource = "SELECT tblnews.* FROM tblnews where username = '" & main.SQLCon.Recordset![UserName] & "';"
    
    main.SQLCon2.Refresh
    lngCounter = 0
    Do While Not main.SQLCon2.Recordset.EOF
        'delete the record
        main.SQLCon2.Recordset.Delete
        main.SQLCon2.Recordset.MoveNext
        lngCounter = lngCounter + 1
    Loop
    
    Print #1, "All news deleted.  " & vbCrLf & vbCrLf
    
    Printer.Print "News Deleted: " & lngCounter
    'clean up
    main.SQLCon2.Recordset.Close
    
    'clear out the users attack results
    Print #1, "Clearing out attack results.  " & vbCrLf
    
    main.SQLCon2.RecordSource = "SELECT tblattackresults.* FROM tblattackresults WHERE username = '" & main.SQLCon.Recordset![UserName] & "';"
    
    main.SQLCon2.Refresh
    lngCounter = 0
    Do While Not main.SQLCon2.Recordset.EOF
        'delete the records
        main.SQLCon2.Recordset.Delete
        main.SQLCon2.Recordset.MoveNext
        lngCounter = lngCounter + 1
    Loop
    
    'clean up
    main.SQLCon2.Recordset.Close
    
    Print #1, "All attack reports deleted.  " & vbCrLf & vbCrLf
    
    
    'clear out the intelegance reports for the user
    Print #1, "Clearing out intelegance reports" & vbCrLf
    
    main.SQLCon2.RecordSource = "SELECT TblIntelegenceReports.* FROM TblIntelegenceReports WHERE username = '" & main.SQLCon.Recordset![UserName] & "';"
    
    main.SQLCon2.Refresh
    
    Do While Not main.SQLCon2.Recordset.EOF
        'delete the records
        main.SQLCon2.Recordset.Delete
        main.SQLCon2.Recordset.MoveNext
        lngCounter = lngCounter + 1
    Loop
    
    'clean up
    main.SQLCon2.Recordset.Close
    
    Print #1, "Intelegeance reports have been cleared. " & vbCrLf & vbCrLf
    
    Print #1, "Cleaning contient"
    
    main.SQLCon2.RecordSource = "SELECT tblcontinents.fullup, newaccounts, numaccounts FROm tblcontinents WHERE continentno = " & main.SQLCon.Recordset![continent]
    
    main.SQLCon2.Refresh
    
    If Not main.SQLCon2.Recordset.EOF Then
            'update the continent details
            Call main.SQLCon2.Recordset.Update("fullup", 0)
            Call main.SQLCon2.Recordset.Update("newaccounts", 1)
            Call main.SQLCon2.Recordset.Update("numaccounts", CInt(main.SQLCon2.Recordset![NumAccounts]) - 1)
    End If
    
    Print #1, "Content details updated.  " & vbcrfl & vbCrLf
    Printer.Print "Reports Deleted: " & lngCounter
    Printer.Print
    main.SQLCon2.Recordset.Close
    
    Print #1, "Deleteing account" & vbCrLf
    
    main.SQLCon.Recordset.Delete
    Printer.FontName = "CP 860 7.6 CPI"
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print "ACCOUNT DELETED"
    Printer.FontUnderline = False
    Printer.FontBold = False
    Printer.EndDoc
    
    Print #1, "Account Deleted" & vbCrLf & vbCrLf
    
    main.SQLCon.Recordset.MoveNext
Loop
    
        
main.SQLCon.Recordset.Close

Print #1, "Inactives deleted" & vbCrLf & vbCrLf & vbCrLf
    
                    

End Sub
