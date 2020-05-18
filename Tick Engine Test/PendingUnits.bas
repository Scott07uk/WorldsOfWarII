Attribute VB_Name = "PendingUnits"
Sub PendingUnitsProcess()
    Print #1, "+--------------------------+"
    Print #1, "| PROCESSING PENDING UNITS |"
    Print #1, "+--------------------------+"
    
    main.SQLCon.RecordSource = "SELECT tblPendingUnits.* FROM tblPendingUnits;"
    main.SQLCon.Refresh

    Set rspending = main.SQLCon.Recordset
    If Not rspending.EOF Then
        Do While Not rspending.EOF
            rspending![TimeLeft] = rspending![TimeLeft] - 1
            rspending.Update
            Print #1, rspending![UserName] & " has been processed"
            If rspending![TimeLeft] = 0 Then
                main.SQLCon2.RecordSource = "SELECT tblAccounts.Username, numinfantry, numcommandos, numjeeps, nummat1, nummat2, numsherman, numspit1, numspit2, numspit5, numhurr, numlancaster, numwellington, numhalifax, numlandingcraft, numlandingship, numfrigates, numbattleship, numcarrier, numcargoship, numAAguns, numpilboxes, numturrets, numballoons, numwatermines, numMP, numexplorers, numspies FROM tblaccounts WHERE username = '" & rspending![UserName] & "';"
                main.SQLCon2.Refresh
                
                Set rsaccount = main.SQLCon2.Recordset
                If Not rsaccount.EOF Then
                    rsaccount![numinfantry] = rsaccount![numinfantry] + rspending![infantry]
                    rsaccount![numcommandos] = rsaccount![numcommandos] + rspending![comandoes]
                    rsaccount![numjeeps] = rsaccount![numjeeps] + rspending![jeeps]
                    rsaccount![nummat1] = rsaccount![nummat1] + rspending![matilda1]
                    rsaccount![nummat2] = rsaccount![nummat2] + rspending![matilda2]
                    rsaccount![numsherman] = rsaccount![numsherman] + rspending![sherman]
                    rsaccount![numspit1] = rsaccount![numspit1] + rspending![spitfire1]
                    rsaccount![numspit2] = rsaccount![numspit2] + rspending![spitfire2]
                    rsaccount![numspit5] = rsaccount![numspit5] + rspending![spitfire5]
                    rsaccount![numhurr] = rsaccount![numhurr] + rspending![hurricane]
                    rsaccount![numlancaster] = rsaccount![numlancaster] + rspending![lancaster]
                    rsaccount![numwellington] = rsaccount![numwellington] + rspending![wellington]
                    rsaccount![numhalifax] = rsaccount![numhalifax] + rspending![halifax]
                    rsaccount![numlandingcraft] = rsaccount![numlandingcraft] + rspending![landingcraft]
                    rsaccount![numlandingship] = rsaccount![numlandingship] + rspending![landingship]
                    rsaccount![numfrigates] = rsaccount![numfrigates] + rspending![frigates]
                    rsaccount![numbattleship] = rsaccount![numbattleship] + rspending![Battleship]
                    rsaccount![numcarrier] = rsaccount![numcarrier] + rspending![carrier]
                    rsaccount![NumCargoShip] = rsaccount![NumCargoShip] + rspending![CargoShip]
                    rsaccount![numaaguns] = rsaccount![numaaguns] + rspending![AAGuns]
                    rsaccount![numpilboxes] = rsaccount![numpilboxes] + rspending![Pilboxes]
                    rsaccount![numturrets] = rsaccount![numturrets] + rspending![Turrets]
                    rsaccount![numballoons] = rsaccount![numballoons] + rspending![Balloons]
                    rsaccount![numwatermines] = rsaccount![numwatermines] + rspending![mines]
                    rsaccount![NumMP] = rsaccount![NumMP] + rspending![MP]
                    rsaccount![NumExplorers] = rsaccount![NumExplorers] + rspending![Explorers]
                    rsaccount![NumSpies] = rsaccount![NumSpies] + rspending![Spies]
                    
                    rsaccount.Update
                    Print #1, "Units added to accounts with " & rsaccount![UserName]
                    rspending.Delete
                End If
                rsaccount.Close
            End If
        rspending.MoveNext
        Loop
    End If
    rspending.Close
    Set rspending = Nothing
    
    Print #1, "+-------------------------+"
    Print #1, "| PENDING UNITS PROCESSED |"
    Print #1, "+-------------------------+" & vbCrLf & vbCrLf
End Sub

