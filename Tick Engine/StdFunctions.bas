Attribute VB_Name = "StdFunctions"
'declair global varibles
Public strSettingsFile As String       'varible to hold the settings file
Public blnAuditOn As Boolean           'varible to store if auditing is on
Public intTickTimes(60) As Integer     'varible to hold the times of the ticks
Public blnCanTicksStart As Boolean     'varible to decide if the ticks are allowed to start
Public intLoopCounter As Integer       'varible used for loops
Public intNumTicks As Integer          'varible used to decide how many ticks there are per hour
Public intTimerHours As Integer        'varible used for the tick timers
Public intTimerMinutes As Integer      'varible used for the tick timers
Public intTimerSeconds As Integer      'varible used for the tick timers
Public blnUserLogedin As Boolean       'varible used to see if the user is loged in
Public intLogedInAction As Integer     'varible to store what the user is doing loged in
Public strAuditFile As String          'varible to hold the name of the audit file


Public Sub syncTimer()

'set the timers clock hours
intTimerHours = Hour(Now())
intTimerMinutes = Minute(Now())
intTimerSeconds = Second(Now())

End Sub

Public Sub returnProcessing()
Select Case intLogedInAction
Case 1
    'the user was updating the setting
    Open strSettingsFile For Output As #1
    If frmSettings.auditticks.Text = "Yes" Then
            Print #1, True
        Else
            Print #1, False
    End If
    
    Close #1
    
    Unload frmSettings
    main.Enabled = True
    main.SetFocus

Case 2
    main.Enabled = True
    main.SetFocus
    Call processTick
    
Case 3
    main.SQLCon.RecordSource = "SELECT tblticktimes.* FROM tblticktimes WHERE ticktime = " & CInt(FrmAddTime.Ticktime.Text)
    
    main.SQLCon.Refresh
    
    If main.SQLCon.Recordset.EOF Then
            main.SQLCon.Recordset.AddNew
            main.SQLCon.Recordset![Ticktime] = CInt(FrmAddTime.Ticktime.Text)
            main.SQLCon.Recordset.Update
    End If
    
    main.SQLCon.Recordset.Close
    
    main.SQLCon.RecordSource = "SELECT * FROM tblticktimes;"
    
    main.SQLCon.Refresh
    
    If main.SQLCon.Recordset.EOF Then
        blnCanTicksStart = False
        main.lblticktimes.Caption = "NONE SET"
    Else
        blnCanTicksStart = True
        
        'load the tick times
        intLoopCounter = 0
        main.lblticktimes.Caption = ""
        Do While Not main.SQLCon.Recordset.EOF
            intTickTimes(intLoopCounter) = main.SQLCon.Recordset![Ticktime]
            intLoopCounter = intLoopCounter + 1
            main.lblticktimes.Caption = main.lblticktimes.Caption & main.SQLCon.Recordset![Ticktime] & vbCrLf
            
            main.SQLCon.Recordset.MoveNext
        Loop
        
        'set the number of ticks to the number of times the loop was ran
        intNumTicks = intLoopCounter
End If
    main.SQLCon.Recordset.Close
    Unload FrmAddTime
    main.Enabled = True
    main.SetFocus
Case 4
    main.SQLCon.RecordSource = "SELECT * FROM tblticktimes WHERE ticktime = " & CInt(frmRemoveTime.removetime.Text)
    
    main.SQLCon.Refresh
    
    If Not main.SQLCon.Recordset.EOF Then
            main.SQLCon.Recordset.Delete
    End If
    
    main.SQLCon.Recordset.Close
    
    main.SQLCon.RecordSource = "SELECT * from tblticktimes;"
    main.SQLCon.Refresh
        If main.SQLCon.Recordset.EOF Then
        blnCanTicksStart = False
        main.lblticktimes.Caption = "NONE SET"
    Else
        blnCanTicksStart = True
        
        'load the tick times
        intLoopCounter = 0
        main.lblticktimes.Caption = ""
        Do While Not main.SQLCon.Recordset.EOF
            intTickTimes(intLoopCounter) = main.SQLCon.Recordset![Ticktime]
            intLoopCounter = intLoopCounter + 1
            main.lblticktimes.Caption = main.lblticktimes.Caption & main.SQLCon.Recordset![Ticktime] & vbCrLf
            
            main.SQLCon.Recordset.MoveNext
        Loop
        
        'set the number of ticks to the number of times the loop was ran
        intNumTicks = intLoopCounter
End If
main.SQLCon.Recordset.Close
Unload frmRemoveTime
main.Enabled = True
main.SetFocus
    
End Select
            

End Sub

Public Sub cancelProcessing()
Select Case intLogedInAction
Case 1
    Unload frmSettings
    main.Enabled = True
    main.SetFocus
    
Case 2
    main.Enabled = True
    main.SetFocus
Case 3
    Unload FrmAddTime
    main.Enabled = True
    main.SetFocus
Case 4
    Unload frmRemoveTime
    main.Enabled = True
    main.SetFocus
    
End Select
End Sub

Public Sub processTick()
    Dim intTickNo As Integer

    'this will handle the tick
    main.Enabled = False
    frmProcessing.Show
    
    'code to handle the tick goes here
    
    'get the tick number
    main.SQLCon.RecordSource = "SELECT tickno FROM tblticks;"
    
    main.SQLCon.Refresh
    
    intTickNo = CInt(main.SQLCon.Recordset![TickNo]) + 1
    
    main.SQLCon.Recordset![TickNo] = intTickNo
    
    main.SQLCon.Recordset.Update
    
    main.SQLCon.Recordset.Close
    
    'begin auditing
    strAuditFile = App.Path & "\audit\" & intTickNo & ".aud"
    
    Open strAuditFile For Output As #1
    
    Print #1, "+----------------------------------------+" & vbCrLf
    Print #1, "| START WORLDS OF WAR II TICK AUDIT FILE |" & vbCrLf
    Print #1, "+----------------------------------------+" & vbCrLf & vbCrLf
    
    
    Call incLastLogin
    Call resource
    Call emailInactives
    Call delInactiveAccounts
    Call ResearchTicks
    Call ConTicks
    Call emailinactive
    Call PendingBuildingProcess
    Call PendingUnitsProcess
    Call attack
    Call voteValidation
    'end of the tick processing
    
    Close #1
    Unload frmProcessing
    main.Enabled = True
    main.SetFocus
End Sub
