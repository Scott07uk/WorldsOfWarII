Attribute VB_Name = "PendingBuilding"
Sub PendingBuildingProcess()

    main.SQLCon.RecordSource = "SELECT tblpendingbuildings.* FROM tblpendingbuildings;"
    main.SQLCon.Refresh

    Set Rspending = main.SQLCon.Recordset
    If Not Rspending.EOF Then
        Do While Not Rspending.EOF
            Rspending![ticksleft] = Rspending![ticksleft] - 1
            Rspending.Update
            If Rspending![ticksleft] = 0 Then
                main.SQLCon2.RecordSource = "SELECT tblAccounts.Username, NumFarms, NumMines, NumRefineries, NumMills FROM tblaccounts WHERE username = '" & Rspending![UserName] & "';"
                main.SQLCon2.Refresh
                
                Set rsAccount = main.SQLCon2.Recordset
                If Not rsAccount.EOF Then
                    rsAccount![numfarms] = rsAccount![numfarms] + Rspending![farms]
                    rsAccount![nummines] = rsAccount![nummines] + Rspending![mines]
                    rsAccount![numrefineries] = rsAccount![numrefineries] + Rspending![refineries]
                    rsAccount![nummills] = rsAccount![nummills] + Rspending![sawmills]
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
