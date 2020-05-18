Attribute VB_Name = "SysValidation"
Public Sub voteValidation()
Print #1, "+-----------------------------+"
Print #1, "| VALIDATION POWER BASE VOTES |"
Print #1, "+-----------------------------+"
Print #1, ""
Dim intVoteCounter As Integer
Dim intNumAccounts As Integer

intNumAccounts = 0
intVoteCounter = 0

main.SQLCon.RecordSource = "SELECT username, numvotes from tblaccounts;"

main.SQLCon.Refresh

Do While Not main.SQLCon.Recordset.EOF
    main.SQLCon2.RecordSource = "SELECT suspended FROM tblaccounts WHERE vote = '" & main.SQLCon.Recordset![UserName] & "';"
    
    main.SQLCon2.Refresh
    
    If main.SQLCon2.Recordset.EOF Then
            intVoteCounter = 0
        Else
            Do While Not main.SQLCon2.Recordset.EOF
                If CBool(main.SQLCon2.Recordset![Suspended]) = False Then
                        intVoteCounter = intVoteCounter + 1
                End If
                main.SQLCon2.Recordset.MoveNext
            Loop
    End If
    
    If CInt(main.SQLCon.Recordset![NumVotes]) <> intVoteCounter Then
            Print #1, "Error detected and corrected for " & main.SQLCon.Recordset![UserName]
            main.SQLCon.Recordset![NumVotes] = intVoteCounter
            main.SQLCon.Recordset.Update
    End If
    
    intVoteCounter = 0
    
    main.SQLCon2.Recordset.Close
    
    main.SQLCon.Recordset.MoveNext
Loop

Print #1, vbCrLf & vbCrLf
Print #1, "+-------------------------------+"
Print #1, "| VALIDATING CONTINENT SETTINGS |"
Print #1, "+-------------------------------+"
Print #1, ""
main.SQLCon.Recordset.Close

main.SQLCon.RecordSource = "SELECT continentno, newaccounts, numaccounts, locked FROM tblcontinents;"

main.SQLCon.Refresh

Do While Not main.SQLCon.Recordset.EOF
    main.SQLCon2.RecordSource = "SELECT powerbase FROM tblaccounts WHERE continent = " & main.SQLCon.Recordset![ContinentNo] & " ORDER BY numvotes DESC;"
    
    main.SQLCon2.Refresh
    
    If Not main.SQLCon2.Recordset.EOF Then
            main.SQLCon2.Recordset![PowerBase] = 1
            intNumAccounts = 1
            
            main.SQLCon2.Recordset.Update
            
            main.SQLCon2.Recordset.MoveNext
            
            Do While Not main.SQLCon2.Recordset.EOF
                main.SQLCon2.Recordset![PowerBase] = 0
                intNumAccounts = intNumAccounts + 1
                main.SQLCon2.Recordset.MoveNext
            Loop
    End If
    
    main.SQLCon2.Recordset.Close
    
    If CInt(main.SQLCon.Recordset![NumAccounts]) <> intNumAccounts Then
            main.SQLCon.Recordset![NumAccounts] = intNumAccounts
            
            main.SQLCon.Recordset.Update
    End If
    
    If intNumAccounts > 14 Then
            main.SQLCon.Recordset![NewAccounts] = 0
            main.SQLCon.Recordset.Update
        Else
            If CBool(main.SQLCon.Recordset![Locked]) = False Then
                    main.SQLCon.Recordset![NewAccounts] = 1
                    main.SQLCon.Recordset.Update
            End If
    End If
    
    main.SQLCon.Recordset.MoveNext
Loop
            
main.SQLCon.Recordset.Close

Print #1, vbCrLf & vbCrLf
Print #1, "+------------------------+"
Print #1, "| VALIDATING UNIT LEVELS |"
Print #1, "+------------------------+"
Print #1, ""
main.SQLCon.RecordSource = "SELECT infantry, comandoes FROM tblmovements;"

main.SQLCon.Refresh

Do While Not main.SQLCon.Recordset.EOF
    If main.SQLCon.Recordset![infantry] < 0 Then
            main.SQLCon.Recordset![infantry] = 0
    End If
    If main.SQLCon.Recordset![comandoes] < 0 Then
            main.SQLCon.Recordset![comandoes] = 0
    End If
    main.SQLCon.Recordset.Update
    main.SQLCon.Recordset.MoveNext
Loop

main.SQLCon.Recordset.Close

main.SQLCon.RecordSource = "SELECT numinfantry, numcommandos FROM tblaccounts;"

main.SQLCon.Refresh

Do While Not main.SQLCon.Recordset.EOF
    If main.SQLCon.Recordset![numinfantry] < 0 Then
            main.SQLCon.Recordset![numinfantry] = 0
    End If
    If main.SQLCon.Recordset![numcommandos] < 0 Then
            main.SQLCon.Recordset![numcommandos] = 0
    End If
    main.SQLCon.Recordset.Update
    main.SQLCon.Recordset.MoveNext
Loop
main.SQLCon.Recordset.Close

Print #1, "+--------------------------+"
Print #1, "| ALL VALIDATION COMPLEATE |"
Print #1, "+--------------------------+" & vbCrLf & vbCrLf

End Sub
