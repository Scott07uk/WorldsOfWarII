Attribute VB_Name = "PendingUnits"
Sub PendingUnitsProcess()

    main.SQLCon.RecordSource = "SELECT tblPendingUnits.* FROM tblPendingUnits;"
    main.SQLCon.Refresh

    Set Rspending = main.SQLCon.Recordset
    If Not Rspending.EOF Then
        Do While Not Rspending.EOF
            Rspending![TimeLeft] = Rspending![TimeLeft] - 1
            Rspending.Update
            If Rspending![TimeLeft] = 0 Then
                main.SQLCon2.RecordSource = "SELECT tblAccounts.Username, numinfantry, numcommandos, numjeeps, nummat1, nummat2, numsherman, numspit1, numspit2, numspit5, numhurr, numlancaster, numwellington, numhalifax, numlandingcraft, numlandingship, numfrigates, numbattleship, numcarrier, numcargoship, numAAguns, numpilboxes, numturrets, numballoons, numwatermines, numMP, numexplorers, numspies FROM tblaccounts WHERE username = '" & Rspending![UserName] & "';"
                main.SQLCon2.Refresh
                
                Set rsAccount = main.SQLCon2.Recordset
                If Not rsAccount.EOF Then
                    rsAccount![numinfantry] = rsAccount![numinfantry] + Rspending![infantry]
                    rsAccount![numcommandos] = rsAccount![numcommandos] + Rspending![comandoes]
                    rsAccount![numjeeps] = rsAccount![numjeeps] + Rspending![jeeps]
                    rsAccount![nummat1] = rsAccount![nummat1] + Rspending![matilda1]
                    rsAccount![nummat2] = rsAccount![nummat2] + Rspending![matilda2]
                    rsAccount![numsherman] = rsAccount![numsherman] + Rspending![sherman]
                    rsAccount![numspit1] = rsAccount![numspit1] + Rspending![spitfire1]
                    rsAccount![numspit2] = rsAccount![numspit2] + Rspending![spitfire2]
                    rsAccount![numspit5] = rsAccount![numspit5] + Rspending![spitfire5]
                    rsAccount![numhurr] = rsAccount![numhurr] + Rspending![hurricane]
                    rsAccount![NumLancaster] = rsAccount![NumLancaster] + Rspending![lancaster]
                    rsAccount![NumWellington] = rsAccount![NumWellington] + Rspending![wellington]
                    rsAccount![NumHalifax] = rsAccount![NumHalifax] + Rspending![halifax]
                    rsAccount![NumLandingCraft] = rsAccount![NumLandingCraft] + Rspending![LandingCraft]
                    rsAccount![NumLandingShip] = rsAccount![NumLandingShip] + Rspending![LandingShip]
                    rsAccount![numfrigates] = rsAccount![numfrigates] + Rspending![frigates]
                    rsAccount![numbattleship] = rsAccount![numbattleship] + Rspending![Battleship]
                    rsAccount![numcarrier] = rsAccount![numcarrier] + Rspending![carrier]
                    rsAccount![NumCargoShip] = rsAccount![NumCargoShip] + Rspending![CargoShip]
                    rsAccount![numaaguns] = rsAccount![numaaguns] + Rspending![AAGuns]
                    rsAccount![numpilboxes] = rsAccount![numpilboxes] + Rspending![Pilboxes]
                    rsAccount![numturrets] = rsAccount![numturrets] + Rspending![Turrets]
                    rsAccount![numballoons] = rsAccount![numballoons] + Rspending![Balloons]
                    rsAccount![numwatermines] = rsAccount![numwatermines] + Rspending![mines]
                    rsAccount![NumMP] = rsAccount![NumMP] + Rspending![MP]
                    rsAccount![NumExplorers] = rsAccount![NumExplorers] + Rspending![Explorers]
                    rsAccount![NumSpies] = rsAccount![NumSpies] + Rspending![Spies]
                    
                    rsAccount.Update
                    Rspending.Delete
                End If
                rsAccount.Close
            End If
        Rspending.MoveNext
        Loop
    End If
    Rspending.Close
End Sub

