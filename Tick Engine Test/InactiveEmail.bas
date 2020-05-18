Attribute VB_Name = "InactiveEmail"
Public Sub emailinactive()
    
    main.SQLCon.RecordSource = "SELECT tblactiveemails.* FROM tblactiveemails;"
    main.SQLCon.Refresh
    
    Set rsinactive = main.SQLCon.Recordset
    If Not rsinactive.EOF Then
        Do While Not rsinactive.EOF 'loops through all emails that are not active
            If rsinactive![Ticks] > 72 Then
                    rsinactive.Delete 'delete all email records not activated within 72 tick
                Else
                    rsinactive![Ticks] = CInt(rsinactive![Ticks]) + 1
                    rsinactive.Update
            End If
        rsinactive.MoveNext
        Loop
    End If
    rsinactive.Close
End Sub
