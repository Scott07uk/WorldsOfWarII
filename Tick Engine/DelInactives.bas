Attribute VB_Name = "DelInactives"
Public Sub emailInactives()
Dim intNextEmail As Integer
Print #1, "+------------------------------------------------------+" & vbCrLf
Print #1, "| Emailing accounts that have been inactive for 7 days |" & vbCrLf
Print #1, "+------------------------------------------------------+" & vbCrLf & vbCrLf

main.SQLCon.RecordSource = "SELECT email, username FROM tblaccounts WHERE lastlogin = 168 AND holiday = 0"

main.SQLCon.Refresh

main.SQLCon2.RecordSource = "SELECT tblemailqueue.* FROM tblemailqueue order by emailID DESC;"

main.SQLCon2.Refresh

If main.SQLCon2.Recordset.EOF Then
        intNextEmail = 1
    Else
        intNextEmail = CInt(main.SQLCon2.Recordset![EmailID]) + CInt(1)
End If

Do While Not main.SQLCon.Recordset.EOF
    Print #1, main.SQLCon.Recordset![Username] & " has been inactive for 7 days, email has been added to the queue" & vbCrLf
    main.SQLCon2.Recordset.AddNew
    main.SQLCon2.Recordset![EmailID] = intNextEmail
    main.SQLCon2.Recordset![Emailto] = main.SQLCon.Recordset![Email]
    main.SQLCon2.Recordset![EmailSubject] = "Your Worlds of War Account"
    'message to be spell checked
    main.SQLCon2.Recordset![EmailBody] = "<font face=arial size=2>" & main.SQLCon.Recordset![Username] & "<BR>Your account with Worlds of War II has been inactive for 7 days now, If you do not wish you account to be deleted then please login to <A HREF=http://www.worldsofwar.co.uk target=_blank>www.worldsofwar.co.uk</A> and your account will be safe.  If you no longer wish to play Worlds of War II, you need not do anything, your account will be queued for deletion in the next few days.  <BR /><BR /><I>Worlds of War Creators</I> </font>"
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

Print #1, "+-----------------------------------------+" & vbCrLf
Print #1, "| Incrementing last login on all accounts |" & vbCrLf
Print #1, "+-----------------------------------------+" & vbCrLf & vbCrLf

'method to change the last login time
    main.SQLCon.RecordSource = "SELECT tblaccounts.lastlogin, ticksmsg, username, protectiontime FROM tblaccounts;"
    
    main.SQLCon.Refresh
    
    Do While Not main.SQLCon.Recordset.EOF
        Call main.SQLCon.Recordset.Update("lastlogin", CLng(CLng(main.SQLCon.Recordset![LastLogin]) + 1))
        Call main.SQLCon.Recordset.Update("ticksmsg", 0)
        If main.SQLCon.Recordset![ProtectionTime] > 0 Then
                Call main.SQLCon.Recordset.Update("protectiontime", main.SQLCon.Recordset![ProtectionTime] - 1)
        End If
        Print #1, main.SQLCon.Recordset![Username] & " incremented successfully to " & CLng(main.SQLCon.Recordset![LastLogin])
        main.SQLCon.Recordset.MoveNext
    Loop
    
    main.SQLCon.Recordset.Close
        
    Print #1, "All accounts have been incremented" & vbCrLf & vbCrLf

End Sub

Public Sub delInactiveAccounts()

Print #1, "+-----------------------------+" & vbCrLf
Print #1, "| Deleteing inactive accounts |" & vbCrLf
Print #1, "+-----------------------------+" & vbCrLf & vbCrLf

main.SQLCon.RecordSource = "SELECT tblaccounts.* FROM tblaccounts WHERE lastlogin = 240;"

main.SQLCon.Refresh

Do While Not main.SQLCon.Recordset.EOF
    Print #1, main.SQLCon.Recordset![Username] & " is inactive and is being deleted"
    'remove the vote for the power base
    
    'check to see if the person is voting for themselfs
    If main.SQLCon.Recordset![Username] = main.SQLCon.Recordset![Vote] Then
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
    main.SQLCon2.RecordSource = "SELECT * FROM tblmessages WHERE tocontinent = " & main.SQLCon.Recordset![Continent] & " AND tocountry = " & main.SQLCon.Recordset![CountryID]
    
    main.SQLCon2.Refresh
    
    Do While Not main.SQLCon2.Recordset.EOF
        'delete the message
        main.SQLCon2.Recordset.Delete
        'audit the delete
        Print #1, "Message to " & main.SQLCon.Recordset![Username] & " deleted." & vbCrLf
        'mode to next message
        main.SQLCon2.Recordset.MoveNext
    Loop
    'clean up
    main.SQLCon2.Recordset.Close
    'messages to deleted
    Print #1, "All messages to " & main.SQLCon.Recordset![Username] & " deleted." & vbcrfl & vbCrLf
    Print #1, "Deleting messages from " & main.SQLCon.Recordset![Username] & vbCrLf
    
    'clear out the messages from
    main.SQLCon2.RecordSource = "SELECT tblmessages.* FROM tblmessages WHERE fromcontinent = " & main.SQLCon.Recordset![Continent] & " AND fromcountry = " & main.SQLCon.Recordset![CountryID]
    
    main.SQLCon2.Refresh
    
    Do While Not main.SQLCon2.Recordset.EOF
        'delete message
        main.SQLCon2.Recordset.Delete
        'audit the delete
        Print #1, "Message from " & main.SQLCon.Recordset![Username] & " deleted." & vbCrLf
        'move to next one
        main.SQLCon2.Recordset.MoveNext
    Loop
    
    'clear up
    main.SQLCon2.Recordset.Close
    Print #1, "All messages from " & main.SQLCon.Recordset![Username] & " deleted." & vbCrLf & vbCrLf
    
    
    'clear out all the news for this person
    Print #1, "Clearing out all news items.  " & vbCrLf
    main.SQLCon2.RecordSource = "SELECT tblnews.* FROM tblnews where username = '" & main.SQLCon.Recordset![Username] & "';"
    
    main.SQLCon2.Refresh
    
    Do While Not main.SQLCon2.Recordset.EOF
        'delete the record
        main.SQLCon2.Recordset.Delete
        main.SQLCon2.Recordset.MoveNext
    Loop
    
    Print #1, "All news deleted.  " & vbCrLf & vbCrLf
    
    'clean up
    main.SQLCon2.Recordset.Close
    
    'clear out the users attack results
    Print #1, "Clearing out attack results.  " & vbCrLf
    
    main.SQLCon2.RecordSource = "SELECT tblattackresults.* FROM tblattackresults WHERE username = '" & main.SQLCon.Recordset![Username] & "';"
    
    main.SQLCon2.Refresh
    
    Do While Not main.SQLCon2.Recordset.EOF
        'delete the records
        main.SQLCon2.Recordset.Delete
        main.SQLCon2.Recordset.MoveNext
    Loop
    
    'clean up
    main.SQLCon2.Recordset.Close
    
    Print #1, "All attack reports deleted.  " & vbCrLf & vbCrLf
    
    
    'clear out the intelegance reports for the user
    Print #1, "Clearing out intelegance reports" & vbCrLf
    
    main.SQLCon2.RecordSource = "SELECT TblIntelegenceReports.* FROM TblIntelegenceReports WHERE username = '" & main.SQLCon.Recordset![Username] & "';"
    
    main.SQLCon2.Refresh
    
    Do While Not main.SQLCon2.Recordset.EOF
        'delete the records
        main.SQLCon2.Recordset.Delete
        main.SQLCon2.Recordset.MoveNext
    Loop
    
    'clean up
    main.SQLCon2.Recordset.Close
    
    Print #1, "Intelegeance reports have been cleared. " & vbCrLf & vbCrLf
    
    Print #1, "Cleaning contient"
    
    main.SQLCon2.RecordSource = "SELECT tblcontinents.fullup, newaccounts, numaccounts FROm tblcontinents WHERE continentno = " & main.SQLCon.Recordset![Continent]
    
    main.SQLCon2.Refresh
    
    If Not main.SQLCon2.Recordset.EOF Then
            'update the continent details
            Call main.SQLCon2.Recordset.Update("fullup", 0)
            Call main.SQLCon2.Recordset.Update("newaccounts", 1)
            Call main.SQLCon2.Recordset.Update("numaccounts", CInt(main.SQLCon2.Recordset![NumAccounts]) - 1)
    End If
    
    Print #1, "Content details updated.  " & vbcrfl & vbCrLf
    
    main.SQLCon2.Recordset.Close
    
    Print #1, "Deleteing account" & vbCrLf
    
    main.SQLCon.Recordset.Delete
    
    Print #1, "Account Deleted" & vbCrLf & vbCrLf
    
    main.SQLCon.Recordset.MoveNext
Loop
    
        
main.SQLCon.Recordset.Close

Print #1, "Inactives deleted" & vbCrLf & vbCrLf & vbCrLf
    
                    

End Sub
