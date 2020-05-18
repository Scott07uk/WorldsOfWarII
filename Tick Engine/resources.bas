Attribute VB_Name = "resources"
Sub resource()
    main.SQLCon.RecordSource = "SELECT tblaccounts.food, wood, iron, money, numfarms, nummills, numrefineries, nummines FROM tblaccounts"
    main.SQLCon.Refresh
    Set rsresources = main.SQLCon.Recordset
    Do While Not rsresources.EOF
        Food = 300 + (rsresources![NumFarms] * 100)
        Wood = 300 + (rsresources![NumMills] * 100)
        Iron = 300 + (rsresources![NumMines] * 100)
        Money = 300 + (rsresources![NumRefineries] * 100)
        rsresources![Food] = rsresources![Food] + Food
        rsresources![Wood] = rsresources![Wood] + Wood
        rsresources![Iron] = rsresources![Iron] + Iron
        rsresources![Money] = rsresources![Money] + Money
        rsresources.Update
        rsresources.MoveNext
    Loop
    rsresources.Close
End Sub
