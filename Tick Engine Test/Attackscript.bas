Attribute VB_Name = "Attackscript"
Option Explicit

Sub attack()
    Dim resultno As Long
    Dim attloss As Long
    Dim defloss As Long
    Dim MinesLost As Long
    Dim FarmsLost As Long
    Dim MillsLost As Long
    Dim RefineriesLost As Long
    Dim ratio As Double
    Dim rsattack
    Dim attacknumber As Long
    Dim rsdefend
    Dim intFreeLand As Long
    
    Dim attspitfire1 As Long
    Dim attspitfire2 As Long
    Dim attspitfire5 As Long
    Dim atthurricane As Long
    Dim attlancaster As Long
    Dim attwellington As Long
    Dim atthalifax As Long
    Dim attspitfire1loss As Long
    Dim attspitfire2loss As Long
    Dim attspitfire5loss As Long
    Dim atthurricaneloss As Long
    Dim attlancasterloss As Long
    Dim attwellingtonloss As Long
    Dim atthalifaxloss As Long
    Dim totallost As Double
    Dim attSendArr(18, 25) As Long
    Dim attUsernameArr(25) As String
    Dim defSendArr(23, 25) As Long
    Dim defUsernameArr(25) As String
    Dim attIDArr(25) As Long
    Dim defIDArr(25) As Long
    Dim intAttCounter As Integer
    Dim intDefCounter As Integer
    Dim landGainArr(25) As Integer
    Dim attReportIDArr(25) As Integer
    Dim defReportIDArr(25) As Integer
    
    
    Dim attunittotal As Long
    Dim attforcetotal As Long
    Dim defunittotal As Long
    Dim defforcetotal As Long
    
    Dim defspitfire1 As Long
    Dim defspitfire2 As Long
    Dim defspitfire5 As Long
    Dim defhurricane As Long
    Dim defaagun As Long
    Dim defturrets As Long
    Dim defpillboxes As Long
    Dim attbombingtotal As Long
    Dim turretslost As Long
    Dim pillboxeslost As Long
    Dim num
    Dim defwatermines As Long
    
    Dim defspitfire1loss As Long
    Dim defspitfire2loss As Long
    Dim defspitfire5loss As Long
    Dim defhurricaneloss As Long
    Dim defaagunsloss As Long
    
    Dim amountdefenceland As Long
    Dim numshiplost As Long
    Dim nummineslost As Long
    
    Dim defcommandoesforce As Long
    Dim defmatilda1force As Long
    Dim defmatilda2force As Long
    Dim defshermanforce As Long
    Dim defbattleshipsforce As Long
    Dim defcarrierforce As Long
    Dim defspitfire2force As Long
    Dim defspitfire5force As Long
    Dim defhurricaneforce As Long
    
    Dim attcommandoesforce As Long
    Dim attmatilda1force As Long
    Dim attmatilda2force As Long
    Dim attshermanforce As Long
    Dim attbattleshipsforce As Long
    Dim attcarrierforce As Long
    Dim attspitfire2force As Long
    Dim attspitfire5force As Long
    Dim atthurricaneforce As Long
    
    Dim definfantry As Long
    Dim defcommandoes As Long
    Dim defjeeps As Long
    Dim defmatilda1 As Long
    Dim defmatilda2 As Long
    Dim defsherman As Long
    Dim defLandingcraft As Long
    Dim defLandingship As Long
    Dim deffrigates As Long
    Dim defbattleships As Long
    Dim defcarrier As Long
    Dim defbarrageballoons As Long
                  
    Dim attinfantry As Long
    Dim attcommandoes As Long
    Dim attjeeps As Long
    Dim attmatilda1 As Long
    Dim attmatilda2 As Long
    Dim attsherman As Long
    Dim attLandingcraft As Long
    Dim attLandingship As Long
    Dim attfrigates As Long
    Dim attbattleships As Long
    Dim attcarrier As Long
    
    Dim definfantryloss As Long
    Dim defcommandoesloss As Long
    Dim defjeepsloss As Long
    Dim defmatilda1loss As Long
    Dim defmatilda2loss As Long
    Dim defshermanloss As Long
    Dim defLandingcraftloss As Long
    Dim defLandingshiploss As Long
    Dim deffrigatesloss As Long
    Dim defbattleshipsloss As Long
    Dim defcarrierloss As Long
    Dim defbarrageballoonsloss As Long
    
    Dim attinfantryloss As Long
    Dim attcommandoesloss As Long
    Dim attjeepsloss As Long
    Dim attmatilda1loss As Long
    Dim attmatilda2loss As Long
    Dim attshermanloss As Long
    Dim attLandingcraftloss As Long
    Dim attLandingshiploss As Long
    Dim attfrigatesloss As Long
    Dim attbattleshipsloss As Long
    Dim attcarrierloss As Long
    Dim defencenumber As Integer
    Dim rsresults
    Dim rsreturn
    Dim rsaccount
    Dim index As Integer
    Dim inttocontinent As Integer
    Dim inttocountry As Integer
    Dim defBuildLoss As Double
    Dim intPendingFarms As Integer
    Dim intPendingMines As Integer
    Dim intPendingRefineries As Integer
    Dim intPendingMills As Integer
    Dim intPendingFarmsLost As Integer
    Dim intPendingMinesLost As Integer
    Dim intPendingRefineriesLost As Integer
    Dim intPendingMillsLost As Integer
    
    
    
    main.SQLCon.RecordSource = "SELECT ETA, orderno FROM tblmovements Order BY ORderno ASC;"
    main.SQLCon.Refresh
    
    Set rsattack = main.SQLCon.Recordset
    On Error Resume Next
    If Not rsattack.EOF Then
        Do While Not rsattack.EOF
            rsattack![eta] = rsattack![eta] - 1
            rsattack.Update
            main.SQLCon.Recordset.Update
            rsattack.MoveNext
        Loop
    End If
    main.SQLCon2.Recordset.LockType = 3
    main.SQLCon2.Recordset.CursorType = 2
    
    main.SQLCon2.RecordSource = "SELECT numinfantry, numcommandos, numjeeps, nummat1, nummat2, numsherman, numspit1, numspit2, numspit5, numhurr, numlandingcraft, numlandingship, numlancaster, numhalifax, numwellington, numfrigates, numbattleship, numcarrier, amountland, numfarms, nummines, numrefineries, nummills, numaaguns, numturrets, numballoons, numwatermines, continent, countryid, numpilboxes, username FROM tblaccounts;"
    main.SQLCon2.Refresh
    
    Do While Not main.SQLCon2.Recordset.EOF
        attacknumber = -1
        defencenumber = -1
        intAttCounter = -1
        intDefCounter = -1
        defBuildLoss = 0
        intPendingFarms = 0
        intPendingMines = 0
        intPendingRefineries = 0
        intPendingMills = 0
        intPendingFarmsLost = 0
        intPendingMinesLost = 0
        intPendingRefineriesLost = 0
        intPendingMillsLost = 0
        amountdefenceland = 0
        
        main.SQLCon.RecordSource = "SELECT tblmovements.* FROM tblmovements WHERE ETA <= 0 AND tocontinent = " & main.SQLCon2.Recordset![continent] & " AND tocountry = " & main.SQLCon2.Recordset![countryid]
    
        main.SQLCon.Refresh
    
        Set rsattack = main.SQLCon.Recordset
        If Not rsattack.EOF Then
            Do While Not rsattack.EOF
                If rsattack![status] = 1 Then
                        attacknumber = attacknumber + 1
                        intAttCounter = intAttCounter + 1
                        Print #1, "+----------------------------------------+"
                        Print #1, "| ATTACK DATA " & attacknumber & " |"
                        Print #1, "+----------------------------------------+" & vbCrLf
        
                        attspitfire1 = attspitfire1 + rsattack![spitfire1]
                        attSendArr(6, attacknumber) = rsattack![spitfire1]
                        attspitfire2 = attspitfire2 + rsattack![spitfire2]
                        attSendArr(7, attacknumber) = rsattack![spitfire2]
                        attspitfire5 = attspitfire5 + rsattack![spitfire5]
                        attSendArr(8, attacknumber) = rsattack![spitfire5]
                        atthurricane = atthurricane + rsattack![hurricane]
                        attSendArr(9, attacknumber) = rsattack![hurricane]
                        attlancaster = attlancaster + rsattack![lancaster]
                        attSendArr(12, attacknumber) = rsattack![lancaster]
                        attwellington = attwellington + rsattack![wellington]
                        attSendArr(11, attacknumber) = rsattack![wellington]
                        atthalifax = atthalifax + rsattack![halifax]
                        attSendArr(10, attacknumber) = rsattack![halifax]
                        attinfantry = attinfantry + rsattack![infantry]
                        attSendArr(0, attacknumber) = rsattack![infantry]
                        attcommandoes = attcommandoes + rsattack![comandoes]
                        attSendArr(1, attacknumber) = rsattack![comandoes]
                        attjeeps = attjeeps + rsattack![jeeps]
                        attSendArr(2, attacknumber) = rsattack![jeeps]
                        attmatilda1 = attmatilda1 + rsattack![matilda1]
                        attSendArr(3, attacknumber) = rsattack![matilda1]
                        attmatilda2 = attmatilda2 + rsattack![matilda2]
                        attSendArr(4, attacknumber) = rsattack![matilda2]
                        attsherman = attsherman + rsattack![sherman]
                        attSendArr(5, attacknumber) = rsattack![sherman]
                        attLandingcraft = rsattack![landingcraft]
                        attSendArr(13, attacknumber) = rsattack![landingcraft]
                        attLandingship = rsattack![landingship]
                        attSendArr(14, attacknumber) = rsattack![landingship]
                        attfrigates = attfrigates + rsattack![frigates]
                        attSendArr(15, attacknumber) = rsattack![frigates]
                        attbattleships = attbattleships + rsattack![battleships]
                        attSendArr(16, attacknumber) = rsattack![battleships]
                        attcarrier = attcarrier + rsattack![carrier]
                        attSendArr(17, attacknumber) = rsattack![carrier]
                        
                        attUsernameArr(attacknumber) = rsattack![UserName]
                        attIDArr(attacknumber) = rsattack![orderno]
                    
                    ElseIf rsattack![status] = 2 Then
                        intDefCounter = intDefCounter + 1
                        defencenumber = defencenumber + 1
                
                        Print #1, "+----------------------------------------+"
                        Print #1, "| DEFENCE DATA " & defencenumber & " |"
                        Print #1, "+----------------------------------------+" & vbCrLf
                            
                        defspitfire1 = defspitfire1 + rsattack![spitfire1]
                        defSendArr(6, defencenumber) = rsattack![spitfire1]
                        defspitfire2 = defspitfire2 + rsattack![spitfire2]
                        defSendArr(7, defencenumber) = rsattack![spitfire2]
                        defspitfire5 = defspitfire5 + rsattack![spitfire5]
                        defSendArr(8, defencenumber) = rsattack![spitfire5]
                        defhurricane = defhurricane + rsattack![hurricane]
                        defSendArr(9, defencenumber) = rsattack![hurricane]
                        definfantry = definfantry + rsattack![infantry]
                        defSendArr(0, defencenumber) = rsattack![infantry]
                        defcommandoes = defcommandoes + rsattack![comandoes]
                        defSendArr(1, defencenumber) = rsattack![comandoes]
                        defjeeps = defjeeps + rsattack![jeeps]
                        defSendArr(2, defencenumber) = rsattack![jeeps]
                        defmatilda1 = defmatilda1 + rsattack![matilda1]
                        defSendArr(3, defencenumber) = rsattack![matilda1]
                        defmatilda2 = defmatilda2 + rsattack![matilda2]
                        defSendArr(4, defencenumber) = rsattack![matilda2]
                        defsherman = defsherman + rsattack![sherman]
                        defSendArr(5, defencenumber) = rsattack![sherman]
                        defLandingcraft = defLandingcraft + rsattack![landingcraft]
                        defSendArr(13, defencenumber) = rsattack![landingcraft]
                        defLandingship = defLandingship + rsattack![landingship]
                        defSendArr(14, defencenumber) = rsattack![landingship]
                        deffrigates = deffrigates + rsattack![frigates]
                        defSendArr(15, defencenumber) = rsattack![frigates]
                        defbattleships = defbattleships + rsattack![battleships]
                        defSendArr(16, defencenumber) = rsattack![battleships]
                        defcarrier = defcarrier + rsattack![carrier]
                        defSendArr(17, defencenumber) = rsattack![carrier]
                        
                        defUsernameArr(defencenumber) = rsattack![UserName]
                        defIDArr(defencenumber) = rsattack![orderno]
                End If
            'all the sent forces compiled into arrays
            
            'add the home defence into the defending units compile
            rsattack.MoveNext
            Loop
        rsattack.Close
        End If
        If intAttCounter <> -1 Then
        
        intDefCounter = intDefCounter + 1
        defencenumber = defencenumber + 1
        Print #1, "+----------------------------------------+"
        Print #1, "| DEFENCE DATA " & defencenumber & " |"
        Print #1, "+----------------------------------------+" & vbCrLf
        
        defspitfire1 = defspitfire1 + main.SQLCon2.Recordset![numspit1]
        defSendArr(6, defencenumber) = main.SQLCon2.Recordset![numspit1]
        defspitfire2 = defspitfire2 + main.SQLCon2.Recordset![numspit2]
        defSendArr(7, defencenumber) = main.SQLCon2.Recordset![numspit2]
        defspitfire5 = defspitfire5 + main.SQLCon2.Recordset![numspit5]
        defSendArr(8, defencenumber) = main.SQLCon2.Recordset![numspit5]
        defhurricane = defhurricane + main.SQLCon2.Recordset![numhurr]
        defSendArr(9, defencenumber) = main.SQLCon2.Recordset![numhurr]
        definfantry = definfantry + main.SQLCon2.Recordset![numinfantry]
        defSendArr(0, defencenumber) = main.SQLCon2.Recordset![numinfantry]
        defcommandoes = defcommandoes + main.SQLCon2.Recordset![numcommandos]
        defSendArr(1, defencenumber) = main.SQLCon2.Recordset![numcommandos]
        defjeeps = defjeeps + main.SQLCon2.Recordset![numjeeps]
        defSendArr(2, defencenumber) = main.SQLCon2.Recordset![numjeeps]
        defmatilda1 = defmatilda1 + main.SQLCon2.Recordset![nummat1]
        defSendArr(3, defencenumber) = main.SQLCon2.Recordset![nummat1]
        defmatilda2 = defmatilda2 + main.SQLCon2.Recordset![nummat2]
        defSendArr(4, defencenumber) = main.SQLCon2.Recordset![nummat2]
        defsherman = defsherman + main.SQLCon2.Recordset![numsherman]
        defSendArr(5, defencenumber) = main.SQLCon2.Recordset![numsherman]
        defLandingcraft = main.SQLCon2.Recordset![numlandingcraft]
        defSendArr(13, defencenumber) = main.SQLCon2.Recordset![numlandingcraft]
        defLandingship = main.SQLCon2.Recordset![numlandingship]
        defSendArr(14, defencenumber) = main.SQLCon2.Recordset![numlandingship]
        deffrigates = deffrigates + main.SQLCon2.Recordset![numfrigates]
        defSendArr(15, defencenumber) = main.SQLCon2.Recordset![numfrigates]
        defbattleships = defbattleships + main.SQLCon2.Recordset![numbattleship]
        defSendArr(16, defencenumber) = main.SQLCon2.Recordset![numbattleship]
        defcarrier = defcarrier + main.SQLCon2.Recordset![numcarrier]
        defSendArr(17, defencenumber) = main.SQLCon2.Recordset![numcarrier]
        defpillboxes = main.SQLCon2.Recordset![numpilboxes]
        defturrets = main.SQLCon2.Recordset![numturrets]
        defbarrageballoons = main.SQLCon2.Recordset![numballoons]
        defwatermines = main.SQLCon2.Recordset![numwatermines]
        defaagun = main.SQLCon2.Recordset![numaaguns]
                    
        defUsernameArr(defencenumber) = main.SQLCon2.Recordset![UserName]
        
        'deanos code goes below here
        attunittotal = attspitfire1 + attspitfire2 + attspitfire5 + atthurricane
                
        attspitfire2force = attspitfire2 * 2
        attspitfire5force = attspitfire5 * 3
        atthurricaneforce = atthurricane * 2
                
        attforcetotal = attspitfire1 + attspitfire2force + attspitfire5force + atthurricaneforce
                
        'PLANES VS PLANES
                    
        defunittotal = defspitfire1 + defspitfire2 + defspitfire5 + defhurricane
                    
        defspitfire2force = defspitfire2 * 2
        defspitfire5force = defspitfire5 * 3
        defhurricaneforce = defhurricane * 2
                    
        defforcetotal = defspitfire1 + defspitfire2force + defspitfire5force + defhurricaneforce
                    
        If defunittotal > 0 Or attunittotal > 0 Then
                If attforcetotal > 0 And defforcetotal > 0 Then
                        ratio = CDbl(attforcetotal / defforcetotal)
                    Else
                        ratio = 0
                End If
                defloss = ((attforcetotal + defforcetotal) / 5) * ratio
                        
                If attunittotal > 0 And defunittotal > 0 Then
                        ratio = defforcetotal / attforcetotal
                    Else
                        ratio = 0
                End If
                attloss = ((attforcetotal + defforcetotal) / 5) * ratio
                        
                If defloss > defunittotal Then
                        defloss = defunittotal
                End If
                        
                If attloss > attunittotal Then
                        attloss = attunittotal
                End If
                        
                Do While defloss > 0 'Defence loss
                    Randomize
                    num = CInt(Rnd * 8)
                    Select Case num
                        Case 1 To 3
                            If defspitfire1 > 0 Then
                                    defspitfire1 = defspitfire1 - 1
                                    defspitfire1loss = defspitfire1loss + 1
                                    defloss = defloss - 1
                            End If
                        Case 4 To 5
                            If defspitfire2 > 0 Then
                                    defspitfire2 = defspitfire2 - 1
                                    defspitfire2loss = defspitfire2loss + 1
                                    defloss = defloss - 1
                            End If
                        Case 6
                            If defspitfire5 > 0 Then
                                    defspitfire5 = defspitfire5 - 1
                                    defspitfire5loss = defspitfire5loss + 1
                                    defloss = defloss - 1
                            End If
                        Case 7 To 8
                            If defhurricane > 0 Then
                                    defhurricane = defhurricane - 1
                                    defhurricaneloss = defhurricaneloss + 1
                                    defloss = defloss - 1
                            End If
                    End Select
                Loop
                        
                Do While attloss > 0 'Attack loss
                    Randomize
                    num = (Rnd * 8)
                    Select Case CInt(num)
                        Case 1 To 3
                            If attspitfire1 > 0 Then
                                    attspitfire1 = attspitfire1 - 1
                                    attspitfire1loss = attspitfire1loss + 1
                                    attloss = attloss - 1
                            End If
                        Case 4 To 5
                            If attspitfire2 > 0 Then
                                    attspitfire2 = attspitfire2 - 1
                                    attspitfire2loss = attspitfire2loss + 1
                                    attloss = attloss - 1
                            End If
                        Case 6
                            If attspitfire5 > 0 Then
                                    attspitfire5 = attspitfire5 - 1
                                    attspitfire5loss = attspitfire5loss + 1
                                    attloss = attloss - 1
                            End If
                        Case 7 To 8
                            If atthurricane > 0 Then
                                    atthurricane = atthurricane - 1
                                    atthurricaneloss = atthurricaneloss + 1
                                    attloss = attloss - 1
                            End If
                    End Select
                Loop
                        
                Print #1, "ratio " & ratio
                Print #1, "deftotal " & defunittotal
                Print #1, "atttotal " & attunittotal
                Print #1, "attloss " & attloss
                Print #1, "defloss " & defloss
                        
                'BOMBERS VS AAGUNS
                If (defaagun > 0) Or (defhurricane > 0) Then
                        defunittotal = defhurricane + defspitfire1 + defspitfire2 + defspitfire5
                        defforcetotal = defaagun + (defhurricane * 3) + defspitfire1 + defspitfire2 + (defspitfire5 * 2)
                        attunittotal = atthalifax + attlancaster + attwellington
                        attforcetotal = atthalifax + attlancaster + attwellington
                            
                        ratio = attforcetotal / defforcetotal
                        defloss = ((attforcetotal + defforcetotal) / 5) * ratio
                            
                        If attforcetotal > 0 And defforcetotal > 0 Then
                                ratio = defforcetotal / attforcetotal
                            Else
                                ratio = 0
                        End If
                        attloss = ((attforcetotal + defforcetotal) / 5) * ratio
                            
                        If defloss > defunittotal Then
                                defloss = defunittotal
                        End If
                            
                        If attloss > attunittotal Then
                                attloss = attunittotal
                        End If
                            
                        defaagunsloss = defloss
                            
                        Do While attloss > 0 'Attack loss
                            Randomize
                            num = (Rnd * 6)
                            Select Case CInt(num)
                                Case 1 To 2
                                    If attlancaster > 0 Then
                                            attlancaster = attlancaster - 1
                                            attlancasterloss = attlancasterloss + 1
                                            attloss = attloss - 1
                                    End If
                                Case 3 To 4
                                    If atthalifax > 0 Then
                                            atthalifax = atthalifax - 1
                                            atthalifaxloss = atthalifaxloss + 1
                                            attloss = attloss - 1
                                    End If
                                Case 5 To 6
                                    If attwellington > 0 Then
                                            attwellington = attwellington - 1
                                            attwellingtonloss = attwellingtonloss + 1
                                            attloss = attloss - 1
                                    End If
                            End Select
                        Loop
                        
                        Do While defloss > 0 'defence losses
                            Randomize
                            num = (Rnd * 11)
                            Select Case CInt(num)
                                Case 1 To 2
                                    If defhurricane > 0 Then
                                            defhurricane = defhurricane - 1
                                            defhurricaneloss = defhurricaneloss + 1
                                            defloss = defloss - 1
                                    End If
                                Case 3 To 5
                                    If defspitfire1 > 0 Then
                                            defspitfire1 = defspitfire1 - 1
                                            defspitfire1loss = defspitfire1loss + 1
                                            defloss = defloss - 1
                                    End If
                                Case 6 To 8
                                    If defspitfire2 > 0 Then
                                            defspitfire2 = defspitfire2 - 1
                                            defspitfire2loss = defspitfire2loss + 1
                                            defloss = defloss - 1
                                    End If
                                Case 9 To 11
                                    If defspitfire5 > 0 Then
                                            defspitfire5 = defspitfire5 - 1
                                            defspitfire5loss = defspitfire5loss + 1
                                            defloss = defloss - 1
                                    End If
                            End Select
                        Loop
                                
                        Print #1, "ratio bomber " & ratio
                        Print #1, "deftotal bomber " & defunittotal
                        Print #1, "atttotal bomber " & attunittotal
                        Print #1, "attloss bomber " & attloss
                        Print #1, "defloss bomber " & defloss
                End If
        End If
                'PLANES VS STUFF THAT CANT DEFEND ITS SELF
                attbombingtotal = attlancaster + atthalifax + attwellington
                attbombingtotal = attbombingtotal / 4
                
                        
                If attbombingtotal > 0 Then
                        If defturrets > 0 Then
                                turretslost = defturrets / (attbombingtotal / 2)
                            Else
                                turretslost = 0
                        End If
                        If turretslost > defturrets Then
                                turretslost = defturrets
                        End If
                        If defpillboxes > 0 And attbombingtotal > 0 Then
                                pillboxeslost = defpillboxes / (attbombingtotal / 2)
                            Else
                                pillboxeslost = 0
                        End If
                        If pillboxeslost > defpillboxes Then
                                pillboxeslost = defpillboxes
                        End If
                End If
                
                defturrets = defturrets - turretslost
                Print #1, main.SQLCon2.Recordset![nummills]
                Print #1, "farms lost" & FarmsLost
                Print #1, "Mills lost" & MillsLost
                Print #1, "mines lost" & MinesLost
                Print #1, "refineries lost" & RefineriesLost
                Print #1, "Turrets lost" & turretslost
                Print #1, "Pillboxes lost" & pillboxeslost
        'mine loss, also rubbish idea
        If defwatermines > 0 And (attLandingship + attLandingcraft) > 0 Then
                amountdefenceland = main.SQLCon2.Recordset![amountland]
                amountdefenceland = Sqr(amountdefenceland)
                amountdefenceland = amountdefenceland * 3.141592654
                amountdefenceland = amountdefenceland / defwatermines
                amountdefenceland = amountdefenceland / 4
                numshiplost = amountdefenceland * (attLandingship + attLandingcraft)
                nummineslost = amountdefenceland * (attLandingship + attLandingcraft)
        End If
                    
        Print #1, "Ship lost " & numshiplost
        Print #1, "Mines lost " & nummineslost
                    
        'DUDES VS DUDES
        defunittotal = definfantry + defcommandoes + defjeeps + defmatilda1 + defmatilda2 + defsherman + deffrigates + defbattleships + defcarrier + defturrets + defpillboxes + defbarrageballoons
                
        defcommandoesforce = defcommandoes * 2
        defmatilda1force = defmatilda1 * 2
        defmatilda2force = defmatilda2 * 3
        defshermanforce = defsherman * 2
        defbattleshipsforce = defbattleships * 2
        defcarrierforce = defcarrier * 2
                    
        defforcetotal = definfantry + defcommandoesforce + defjeeps + defmatilda1force + defmatilda2force + defshermanforce + deffrigates + defbattleshipsforce + defcarrierforce + defturrets + defpillboxes + defbarrageballoons
                    
        attunittotal = attinfantry + attcommandoes + attjeeps + attmatilda1 + attmatilda2 + attsherman + attfrigates + attbattleships + attcarrier
                
        attcommandoesforce = attcommandoes * 2
        attmatilda1force = attmatilda1 * 2
        attmatilda2force = attmatilda2 * 3
        attshermanforce = attsherman * 2
        attbattleshipsforce = attbattleships * 2
        attcarrierforce = attcarrier * 2
                    
        attforcetotal = attinfantry + attcommandoesforce + attjeeps + attmatilda1force + attmatilda2force + attshermanforce + attfrigates + attbattleshipsforce + attcarrierforce
                    
        If defforcetotal > 0 And attforcetotal > 0 Then
                ratio = attforcetotal / defforcetotal
            Else
                ratio = 0
        End If
        defloss = ((attforcetotal + defforcetotal) / 5) * ratio
        
        
        If ratio = 0 Then
                If attforcetotal > 0 Then
                        Randomize
                        defBuildLoss = CDbl((Rnd * 0.1) + 0.3)
                    Else
                        defBuildLoss = 0
                End If
            Else
                defBuildLoss = (attforcetotal / defforcetotal) * ratio
        End If
        
        
        If defforcetotal > 0 And attforcetotal > 0 Then
                ratio = defforcetotal / attforcetotal
            Else
                ratio = 0
        End If
        attloss = ((attforcetotal + defforcetotal) / 5) * ratio
        'this calculates the land that is taken
        If CLng(attloss) = 0 Then
                totallost = 0
            Else
                If CLng(defloss / (attloss * 2)) < totallost Then
                        totallost = CLng((attloss * 2) / defloss) / 5
                End If
        End If
        If defloss > defunittotal Then
                defloss = defunittotal
        End If
                    
        If attloss > attunittotal Then
                attloss = attunittotal
        End If
        
        
        
                    
        Print #1, "ratio " & ratio
        Print #1, "deftotal " & defunittotal
        Print #1, "atttotal " & attunittotal
        Print #1, "attloss " & attloss
        Print #1, "defloss " & defloss
                    
        defturrets = defturrets - turretslost
        defpillboxes = defpillboxes - pillboxeslost
        defloss = defloss - (pillboxeslost + turretslost)
                    
        Do While defloss > 0 'Defence loss
            num = CInt(Rnd * 21)
            Select Case num
                Case 1 To 3
                    If definfantry > 0 Then
                            definfantry = definfantry - 1
                            definfantryloss = definfantryloss + 1
                            defloss = defloss - 1
                    End If
                Case 4 To 5
                    If defcommandoes > 0 Then
                            defcommandoes = defcommandoes - 1
                            defcommandoesloss = defcommandoesloss + 1
                            defloss = defloss - 1
                    End If
                Case 6
                    If defjeeps > 0 Then
                            defjeeps = defjeeps - 1
                            defjeepsloss = defjeepsloss + 1
                            defloss = defloss - 1
                    End If
                Case 7 To 8
                    If defmatilda1 > 0 Then
                            defmatilda1 = defmatilda1 - 1
                            defmatilda1loss = defmatilda1loss + 1
                            defloss = defloss - 1
                    End If
                Case 9 To 11
                    If defmatilda2 > 0 Then
                            defmatilda2 = defmatilda2 - 1
                            defmatilda2loss = defmatilda2loss + 1
                            defloss = defloss - 1
                    End If
                Case 12 To 13
                    If defsherman > 0 Then
                            defsherman = defsherman - 1
                            defshermanloss = defshermanloss + 1
                            defloss = defloss - 1
                    End If
                Case 14
                    If deffrigates > 0 Then
                            deffrigates = deffrigates - 1
                            deffrigatesloss = deffrigatesloss + 1
                            defloss = defloss - 1
                    End If
                Case 15 To 16
                    If defbattleships > 0 Then
                            defbattleships = defbattleships - 1
                            defbattleshipsloss = defbattleshipsloss + 1
                            defloss = defloss - 1
                    End If
                Case 17 To 18
                    If defcarrier > 0 Then
                            defcarrier = defcarrier - 1
                            defcarrierloss = defcarrierloss + 1
                            defloss = defloss - 1
                    End If
                Case 19
                    If defturrets > 0 Then
                            defturrets = defturrets - 1
                            turretslost = turretslost + 1
                            defloss = defloss - 1
                    End If
                Case 20
                    If defpillboxes > 0 Then
                            defpillboxes = defpillboxes - 1
                            pillboxeslost = pillboxeslost + 1
                            defloss = defloss - 1
                    End If
                Case 21
                    If defbarrageballoons > 0 Then
                            defbarrageballoons = defbarrageballoons - 1
                            defbarrageballoonsloss = defbarrageballoonsloss + 1
                            defloss = defloss - 1
                    End If
            End Select
        Loop
                    
        Do While attloss > 0 'Attack loss
            Randomize
            num = (Rnd * 18)
            Select Case CInt(num)
                Case 1 To 3
                    If attinfantry > 0 Then
                            attinfantry = attinfantry - 1
                            attinfantryloss = attinfantryloss + 1
                            attloss = attloss - 1
                    End If
                Case 4 To 5
                    If attcommandoes > 0 Then
                            attcommandoes = attcommandoes - 1
                            attcommandoesloss = attcommandoesloss + 1
                            attloss = attloss - 1
                    End If
                Case 6
                    If attjeeps > 0 Then
                            attjeeps = attjeeps - 1
                            attjeepsloss = attjeepsloss + 1
                            attloss = attloss - 1
                    End If
                Case 7 To 8
                    If attmatilda1 > 0 Then
                            attmatilda1 = attmatilda1 - 1
                            attmatilda1loss = attmatilda1loss + 1
                            attloss = attloss - 1
                    End If
                Case 9 To 11
                    If attmatilda2 > 0 Then
                            attmatilda2 = attmatilda2 - 1
                            attmatilda2loss = attmatilda2loss + 1
                            attloss = attloss - 1
                    End If
                Case 12 To 13
                    If attsherman > 0 Then
                            attsherman = attsherman - 1
                            attshermanloss = attshermanloss + 1
                            attloss = attloss - 1
                    End If
                Case 14
                    If attfrigates > 0 Then
                            attfrigates = attfrigates - 1
                            attfrigatesloss = attfrigatesloss + 1
                            attloss = attloss - 1
                    End If
                Case 15 To 16
                    If attbattleships > 0 Then
                            attbattleships = attbattleships - 1
                            attbattleshipsloss = attbattleshipsloss + 1
                            attloss = attloss - 1
                    End If
                Case 17 To 18
                    If attcarrier > 0 Then
                            attcarrier = attcarrier - 1
                            attcarrierloss = attcarrierloss + 1
                            attloss = attloss - 1
                    End If
            End Select
        Loop
        attbombingtotal = attbombingtotal
        FarmsLost = main.SQLCon2.Recordset![numfarms] * (attbombingtotal / 14)
        MillsLost = main.SQLCon2.Recordset![nummills] * (attbombingtotal / 8)
        MinesLost = main.SQLCon2.Recordset![nummines] * (attbombingtotal / 11)
        RefineriesLost = main.SQLCon2.Recordset![numrefineries] * (attbombingtotal / 9)
                        
        If FarmsLost > main.SQLCon2.Recordset![numfarms] Then
                FarmsLost = CLng(main.SQLCon2.Recordset![numfarms] - (main.SQLCon2.Recordset![numfarms] / 8))
        End If
        If MillsLost > main.SQLCon2.Recordset![nummills] Then
                MillsLost = CLng(main.SQLCon2.Recordset![nummills] - (main.SQLCon2.Recordset![nummills] / 8))
        End If
        If MinesLost > main.SQLCon2.Recordset![nummines] Then
                MinesLost = CLng(main.SQLCon2.Recordset![nummines] - (main.SQLCon2.Recordset![nummines] / 8))
        End If
        If RefineriesLost > main.SQLCon2.Recordset![numrefineries] Then
                RefineriesLost = CLng(main.SQLCon2.Recordset![numrefineries] - (main.SQLCon2.Recordset![numrefineries] / 8))
        End If
        
        'totallost = MinesLost + FarmsLost + MillsLost + RefineriesLost
        Randomize
        num = (Rnd * 0.4) + 0.25
        totallost = defBuildLoss * main.SQLCon2.Recordset![amountland] * num
        
        If totallost > CLng(main.SQLCon2.Recordset![amountland] * 0.45) Then
                totallost = CLng(main.SQLCon2.Recordset![amountland] * 0.45)
        End If
        
        main.SQLCon.RecordSource = "SELECT farms, mines, refineries, sawmills FROM tblpendingbuildings WHERE username = '" & main.SQLCon2.Recordset![UserName] & "';"
        main.SQLCon.Refresh
        
        intPendingFarms = 0
        intPendingMines = 0
        intPendingRefineries = 0
        intPendingMills = 0
        
        
        Do While Not main.SQLCon.Recordset.EOF
            intPendingFarms = intPendingFarms + main.SQLCon.Recordset![farms]
            intPendingMines = intPendingMines + main.SQLCon.Recordset![mines]
            intPendingRefineries = intPendingRefineries + main.SQLCon.Recordset![refineries]
            intPendingMills = intPendingMills + main.SQLCon.Recordset![sawmills]
            
            main.SQLCon.Recordset.MoveNext
        Loop
        
        main.SQLCon.Recordset.Close
        
        intFreeLand = main.SQLCon2.Recordset![amountland] - (intPendingFarms + intPendingMines + intPendingRefineries + intPendingMills + (main.SQLCon2.Recordset![numfarms] - FarmsLost) + (main.SQLCon2.Recordset![nummills] - MillsLost) + (main.SQLCon2.Recordset![nummines] - MinesLost) + (main.SQLCon2.Recordset![numrefineries] - RefineriesLost))
        
        Do While intFreeLand < totallost
            Randomize
            num = (Rnd * 7)
            num = CLng(num)
            
            Select Case num
                Case 1
                    If (main.SQLCon2.Recordset![numfarms] - FarmsLost) > 0 Then
                            FarmsLost = FarmsLost + 1
                            intFreeLand = intFreeLand + 1
                        ElseIf intPendingFarms > intPendingFarmsLost Then
                            intPendingFarmsLost = intPendingFarmsLost + 1
                            intFreeLand = intFreeLand + 1
                    End If
                Case 2 To 3
                    If (main.SQLCon2.Recordset![nummines] - MinesLost) > 0 Then
                            MinesLost = MinesLost + 1
                            intFreeLand = intFreeLand + 1
                        ElseIf intPendingMines > intPendingMinesLost Then
                            intPendingMinesLost = intPendingMinesLost + 1
                            intFreeLand = intFreeLand + 1
                    End If
                Case 4 To 5
                    If (main.SQLCon2.Recordset![numrefineries] - RefineriesLost) > 0 Then
                            RefineriesLost = RefineriesLost + 1
                            intFreeLand = intFreeLand + 1
                        ElseIf intPendingRefineries > intPendingRefineriesLost Then
                                intPendingRefineriesLost = intPendingRefineriesLost + 1
                                intFreeLand = intFreeLand + 1
                    End If
                Case 6 To 7
                    If (main.SQLCon2.Recordset![nummills] - MillsLost) > 0 Then
                            MillsLost = MillsLost + 1
                            intFreeLand = intFreeLand + 1
                        ElseIf intPendingMills > intPendingMillsLost Then
                            intPendingMillsLost = intPendingMillsLost + 1
                            intFreeLand = intFreeLand + 1
                        
                    End If
            End Select
        Loop
        
        main.SQLCon.RecordSource = "SELECT farms, mines, refineries, sawmills FROM tblpendingbuildings WHERE USERNAME = '" & main.SQLCon2.Recordset![UserName] & "' ORDER BY ticksleft DESC;"
        
        main.SQLCon.Refresh
        
        Do While ((intPendingFarmsLost + intPendingMinesLost + intPendingMillsLost + intPendingRefineriesLost) > 0) And (Not main.SQLCon.Recordset.EOF)
            If main.SQLCon.Recordset![farms] > 0 Then
                    If main.SQLCon.Recordset![farms] > intPendingFarmsLost Then
                            main.SQLCon.Recordset![farms] = main.SQLCon.Recordset![farms] - intPendingFarmsLost
                        Else
                            intPendingFarmsLost = intPendingFarmsLost - main.SQLCon.Recordset![farms]
                            main.SQLCon.Recordset![farms] = 0
                    End If
            End If
            If main.SQLCon.Recordset![mines] > 0 Then
                    If main.SQLCon.Recordset![mines] > intPendingMinesLost Then
                            main.SQLCon.Recordset![mines] = main.SQLCon.Recordset![mines] - intPendingMinesLost
                        Else
                            intPendingMinesLost = intPendingMinesLost - main.SQLCon.Recordset![mines]
                            main.SQLCon.Recordset![mines] = 0
                    End If
            End If
            If main.SQLCon.Recordset![sawmills] > 0 Then
                    If main.SQLCon.Recordset![sawmills] > intPendingMillsLost Then
                            main.SQLCon.Recordset![sawmills] = main.SQLCon.Recordset![sawmills] - intPendingMillsLost
                        Else
                            intPendingMillsLost = intPendingMillsLost - main.SQLCon.Recordset![sawmills]
                            main.SQLCon.Recordset![sawmills] = 0
                    End If
            End If
            If main.SQLCon.Recordset![refineries] > 0 Then
                    If main.SQLCon.Recordset![refineries] > intPendingRefineriesLost Then
                            main.SQLCon.Recordset![refineries] = main.SQLCon.Recordset![refineries] - intPendingRefineriesLost
                        Else
                            intPendingRefineriesLost = intPendingRefineriesLost - main.SQLCon.Recordset![refineries]
                            main.SQLCon.Recordset![refineries] = 0
                    End If
            End If
            If Not main.SQLCon.Recordset.EOF Then
                    main.SQLCon.Recordset.MoveNext
            End If
        Loop
        
        main.SQLCon.Recordset.Close
                
                    
        main.SQLCon.RecordSource = "SELECT tblattackResults.* FROM tblattackresults ORDER BY reportID DESC;"
        main.SQLCon.Refresh
                    
        Set rsresults = main.SQLCon.Recordset
        If rsresults.EOF Then
                resultno = 0
            Else
                resultno = rsresults![reportID]
        End If
        For index = 0 To intAttCounter
            resultno = resultno + 1
            attReportIDArr(index) = resultno
            rsresults.AddNew
            rsresults![reportID] = resultno
            rsresults![countryid] = main.SQLCon2.Recordset![countryid]
            rsresults![ContinetID] = main.SQLCon2.Recordset![continent]
                
            rsresults![definfantrys] = definfantry + definfantryloss
            rsresults![defcomandoess] = defcommandoes + defcommandoesloss
            rsresults![defjeepss] = defjeeps + defjeepsloss
            rsresults![defmatilda1s] = defmatilda1 + defmatilda1loss
            rsresults![defmatilda2s] = defmatilda2 + defmatilda2loss
            rsresults![defshermans] = defsherman + defshermanloss
            rsresults![defspitfire1s] = defspitfire1 + defspitfire1loss
            rsresults![defspitfire2s] = defspitfire2 + defspitfire2loss
            rsresults![defspitfire5s] = defspitfire5 + defspitfire5loss
            rsresults![defhurricanes] = defhurricane + defhurricaneloss
            rsresults![deffrigatess] = deffrigates + deffrigatesloss
            rsresults![defbattleships] = defbattleships + defbattleshipsloss
            rsresults![defcarriers] = defcarrier + defcarrierloss
            rsresults![defaagunss] = defaagun + defaagunsloss
            rsresults![defpilboxess] = defpillboxes + pillboxeslost
            rsresults![defturretss] = defturrets + turretslost
            rsresults![defballoonss] = defbarrageballoons + defbarrageballoonsloss
            rsresults![defminess] = defwatermines + nummineslost
            
            rsresults![definfantryl] = definfantryloss
            rsresults![defcomandoesl] = defcommandoesloss
            rsresults![defjeepsl] = defjeepsloss
            rsresults![defmatilda1l] = defmatilda1loss
            rsresults![defmatilda2l] = defmatilda2loss
            rsresults![defshermanl] = defshermanloss
            rsresults![defspitfire1l] = defspitfire1loss
            rsresults![defspitfire2l] = defspitfire2loss
            rsresults![defspitfire5l] = defspitfire5loss
            rsresults![defhurricanel] = defhurricaneloss
            rsresults![deffrigatesl] = deffrigatesloss
            rsresults![defbattleshipl] = defbattleshipsloss
            rsresults![defcarrierl] = defcarrierloss
            rsresults![defaagunl] = defaagunsloss
            rsresults![defpilboxesl] = pillboxeslost
            rsresults![defturretsl] = turretslost
            rsresults![defballoonsl] = defbarrageballoonsloss
            rsresults![defminesl] = nummineslost
            rsresults![Mode] = 1
            
            If (attinfantry + attinfantryloss) > 0 Then
                    rsresults![youinfantryl] = attinfantryloss * (attSendArr(0, index) / (attinfantry + attinfantryloss))
                Else
                    rsresults![youinfantryl] = 0
            End If
            If (attcommandoes + attcommandoesloss) > 0 Then
                    rsresults![youcomandoesl] = attcommandoesloss * (attSendArr(1, index) / (attcommandoes + attcommandoesloss))
                Else
                    rsresults![youcomandoesl] = 0
            End If
            If (attjeeps + attjeepsloss) > 0 Then
                    rsresults![youjeepsl] = attjeepsloss * (attSendArr(2, index) / (attjeeps + attjeepsloss))
                Else
                    rsresults![youjeepsl] = 0
            End If
            If (attmatilda1 + attmatilda1loss) > 0 Then
                    rsresults![youmatilda1l] = attmatilda1loss * (attSendArr(3, index) / (attmatilda1 + attmatilda1loss))
                Else
                    rsresults![youmatilda1l] = 0
            End If
            If (attmatilda2 + attmatilda2loss) > 0 Then
                    rsresults![youmatilda2l] = attmatilda2loss * (attSendArr(4, index) / (attmatilda2 + attmatilda2loss))
                Else
                    rsresults![youmatilda2l] = 0
            End If
            If (attsherman + attshermanloss) > 0 Then
                    rsresults![youshermanl] = attshermanloss * (attSendArr(5, index) / (attsherman + attshermanloss))
                Else
                    rsresults![youshermanl] = 0
            End If
            If (attspitfire1 + attspitfire1loss) > 0 Then
                    rsresults![youspitfire1l] = attspitfire1loss * (attSendArr(6, index) / (attspitfire1 + attspitfire1loss))
                Else
                    rsresults![youspitfire1l] = 0
            End If
            If (attspitfire2 + attspitfire5loss) > 0 Then
                    rsresults![youspitfire2l] = attspitfire2loss * (attSendArr(7, index) / (attspitfire2 + attspitfire5loss))
                Else
                    rsresults![youspitfire2l] = 0
            End If
            If (attspitfire5 + attspitfire5loss) > 0 Then
                    rsresults![youspitfire5l] = attspitfire5loss * (attSendArr(8, index) / (attspitfire5 + attspitfire5loss))
                Else
                    rsresults![youspitfire5l] = 0
            End If
            If (atthurricane + atthurricaneloss) > 0 Then
                    rsresults![youhurricanel] = atthurricaneloss * (attSendArr(9, index) / (atthurricane + atthurricaneloss))
                Else
                    rsresults![youhurricanel] = 0
            End If
            If (atthalifax + atthalifaxloss) > 0 Then
                    rsresults![youhalifaxl] = atthalifaxloss * (attSendArr(10, index) / (atthalifax + atthalifaxloss))
                Else
                    rsresults![youhalifaxl] = 0
            End If
            If (attwellington + attwellingtonloss) > 0 Then
                    rsresults![youwellingtonl] = attwellingtonloss * (attSendArr(11, index) / (attwellington + attwellingtonloss))
                Else
                    rsresults![youwellingtonl] = 0
            End If
            If (attlancaster + attlancasterloss) > 0 Then
                    rsresults![youlancasterl] = attlancasterloss * (attSendArr(12, index) / (attlancaster + attlancasterloss))
                Else
                    rsresults![youlancasterl] = 0
            End If
            If (attLandingcraft + attLandingcraftloss) > 0 Then
                    rsresults![youlandingcraftl] = attLandingcraft * (attSendArr(13, index) / (attLandingcraft + attLandingcraftloss))
                Else
                    rsresults![youlandingcraftl] = 0
            End If
            If (attLandingship + attLandingshiploss) > 0 Then
                    rsresults![youlandingshipl] = attLandingshiploss * (attSendArr(14, index) / (attLandingship + attLandingshiploss))
                Else
                    rsresults![youlandingshipl] = 0
            End If
            If (attfrigates + attfrigatesloss) > 0 Then
                    rsresults![youfrigatesl] = attfrigatesloss * (attSendArr(15, index) / (attfrigates + attfrigatesloss))
                Else
                    rsresults![youfrigatesl] = 0
            End If
            If (attbattleships + attbattleshipsloss) > 0 Then
                    rsresults![youbattleshipl] = attbattleshipsloss * (attSendArr(16, index) / (attbattleships + attbattleshipsloss))
                Else
                    rsresults![youbattleshipl] = 0
            End If
            If (attcarrier + attcarrierloss) > 0 Then
                    rsresults![youcarrierl] = attcarrierloss * (attSendArr(17, index) / (attcarrier + attcarrierloss))
                Else
                    rsresults![youcarrierl] = 0
            End If
            rsresults![youaagunl] = 0
            rsresults![youpilboxesl] = 0
            rsresults![youturretl] = 0
            
            rsresults![youballoonl] = 0
            rsresults![youminel] = 0
            
            rsresults![youinfantrys] = attSendArr(0, index)
            rsresults![youcomandoess] = attSendArr(1, index)
            rsresults![youjeepss] = attSendArr(2, index)
            rsresults![youmatilda1s] = attSendArr(3, index)
            rsresults![youmatilda2s] = attSendArr(4, index)
            rsresults![youshermans] = attSendArr(5, index)
            rsresults![youspitfire1s] = attSendArr(6, index)
            rsresults![youspitfire2s] = attSendArr(7, index)
            rsresults![youspitfire5s] = attSendArr(8, index)
            rsresults![youhurricanes] = attSendArr(9, index)
            rsresults![youhalifaxs] = attSendArr(10, index)
            rsresults![youwellingtons] = attSendArr(11, index)
            rsresults![youlancasters] = attSendArr(12, index)
            rsresults![youlandingcrafts] = attSendArr(13, index)
            rsresults![youlandingships] = attSendArr(14, index)
            rsresults![youfrigatess] = attSendArr(15, index)
            rsresults![youbattleships] = attSendArr(16, index)
            rsresults![youcarriers] = attSendArr(17, index)
            rsresults![youaaguns] = 0
            rsresults![youpilboxes] = 0
            rsresults![youturrets] = 0
            rsresults![youballons] = 0
            rsresults![youmines] = 0
            rsresults![UserName] = attUsernameArr(index)
                
            rsresults![FarmsLost] = FarmsLost
            rsresults![MillsLost] = MillsLost
            rsresults![RefineriesLost] = RefineriesLost
            rsresults![MinesLost] = MinesLost
            rsresults![landlost] = totallost
            
            rsresults![attinfantrys] = attinfantry + attinfantryloss
            rsresults![attcomandoess] = attcommandoes + attcommandoesloss
            rsresults![attjeepss] = attjeeps + attjeepsloss
            rsresults![attmatilda1s] = attmatilda1 + attmatilda1loss
            rsresults![attmatilda2s] = attmatilda2 + attmatilda2loss
            rsresults![attshermans] = attsherman + attshermanloss
            rsresults![attspitfire1s] = attspitfire1 + attspitfire1loss
            rsresults![attspitfire2s] = attspitfire2 + attspitfire2loss
            rsresults![attspitfire5s] = attspitfire5 + attspitfire5loss
            rsresults![atthurricanes] = atthurricane + atthurricaneloss
            rsresults![attlancasters] = attlancaster + attlancasterloss
            rsresults![attwellingtons] = attwellington + attwellingtonloss
            rsresults![atthalifaxs] = atthalifax + atthalifaxloss
            rsresults![attfrigates] = attfrigates + attfrigatesloss
            rsresults![attbattleships] = attbattleships + attbattleshipsloss
            rsresults![attcarriers] = attcarrier + attcarrierloss
            
            rsresults![attinfantryl] = attinfantryloss
            rsresults![attcomandoesl] = attcommandoesloss
            rsresults![attjeepsl] = attjeepsloss
            rsresults![attmatilda1] = attmatilda1loss
            rsresults![attmatilda2] = attmatilda2loss
            rsresults![attshermanl] = attshermanloss
            rsresults![attspitfire1l] = attspitfire1loss
            rsresults![attspitfire2l] = attspitfire2loss
            rsresults![attspitfire5l] = attspitfire5loss
            rsresults![atthurricanel] = atthurricaneloss
            rsresults![attlancasterl] = attlancasterloss
            rsresults![attwellingtonl] = attwellingtonloss
            rsresults![atthalifaxl] = atthalifaxloss
            rsresults![attfrigatesl] = attfrigatesloss
            rsresults![attbattleshipl] = attbattleshipsloss
            rsresults![attcarrierl] = attcarrierloss
                        
            rsresults.Update
        Next
        For index = 0 To intDefCounter
            resultno = resultno + 1
            defReportIDArr(index) = resultno
            rsresults.AddNew
            rsresults![reportID] = resultno
            rsresults![countryid] = main.SQLCon2.Recordset![countryid]
            rsresults![ContinetID] = main.SQLCon2.Recordset![continent]
                
            rsresults![definfantrys] = definfantry + definfantryloss
            rsresults![defcomandoess] = defcommandoes + defcommandoesloss
            rsresults![defjeepss] = defjeeps + defjeepsloss
            rsresults![defmatilda1s] = defmatilda1 + defmatilda1loss
            rsresults![defmatilda2s] = defmatilda2 + defmatilda2loss
            rsresults![defshermans] = defsherman + defshermanloss
            rsresults![defspitfire1s] = defspitfire1 + defspitfire1loss
            rsresults![defspitfire2s] = defspitfire2 + defspitfire2loss
            rsresults![defspitfire5s] = defspitfire5 + defspitfire5loss
            rsresults![defhurricanes] = defhurricane + defhurricaneloss
            rsresults![deffrigatess] = deffrigates + deffrigatesloss
            rsresults![defbattleships] = defbattleships + defbattleshipsloss
            rsresults![defcarriers] = defcarrier + defcarrierloss
            rsresults![defaagunss] = defaagun + defaagunsloss
            rsresults![defpilboxess] = defpillboxes + pillboxeslost
            rsresults![defturretss] = defturrets + turretslost
            rsresults![defballoonss] = defbarrageballoons + defbarrageballoonsloss
            rsresults![defminess] = defwatermines + nummineslost
            
            rsresults![definfantryl] = definfantryloss
            rsresults![defcomandoesl] = defcommandoesloss
            rsresults![defjeepsl] = defjeepsloss
            rsresults![defmatilda1l] = defmatilda1loss
            rsresults![defmatilda2l] = defmatilda2loss
            rsresults![defshermanl] = defshermanloss
            rsresults![defspitfire1l] = defspitfire1loss
            rsresults![defspitfire2l] = defspitfire2loss
            rsresults![defspitfire5l] = defspitfire5loss
            rsresults![defhurricanel] = defhurricaneloss
            rsresults![deffrigatesl] = deffrigatesloss
            rsresults![defbattleshipl] = defbattleshipsloss
            rsresults![defcarrierl] = defcarrierloss
            rsresults![defaagunl] = defaagunsloss
            rsresults![defpilboxesl] = pillboxeslost
            rsresults![defturretsl] = turretslost
            rsresults![defballoonsl] = defbarrageballoonsloss
            rsresults![defminesl] = nummineslost
            rsresults![Mode] = 1
            
            If (definfantry + definfantryloss) > 0 Then
                    rsresults![youinfantryl] = definfantryloss * (defSendArr(0, index) / (definfantry + definfantryloss))
                Else
                    rsresults![youinfantryl] = 0
            End If
            If (defcommandoes + defcommandoesloss) > 0 Then
                    rsresults![youcomandoesl] = defcommandoesloss * (defSendArr(1, index) / (defcommandoes + defcommandoesloss))
                Else
                    rsresults![youcomandoesl] = 0
            End If
            If (defjeeps + defjeepsloss) > 0 Then
                    rsresults![youjeepsl] = defjeepsloss * (defSendArr(2, index) / (defjeeps + defjeepsloss))
                Else
                    rsresults![youjeepsl] = 0
            End If
            If (defmatilda1 + defmatilda1loss) > 0 Then
                    rsresults![youmatilda1l] = defmatilda1loss * (defSendArr(3, index) / (defmatilda1 + defmatilda1loss))
                Else
                    rsresults![youmatilda1l] = 0
            End If
            If (defmatilda2 + defmatilda2loss) > 0 Then
                    rsresults![youmatilda2l] = defmatilda2loss * (defSendArr(4, index) / (defmatilda2 + defmatilda2loss))
                Else
                    rsresults![youmatilda2l] = 0
            End If
            If (defsherman + defshermanloss) > 0 Then
                    rsresults![youshermanl] = defshermanloss * (defSendArr(5, index) / (defsherman + defshermanloss))
                Else
                    rsresults![youshermanl] = 0
            End If
            If (defspitfire1 + defspitfire1loss) > 0 Then
                    rsresults![youspitfire1l] = defspitfire1loss * (defSendArr(6, index) / (defspitfire1 + defspitfire1loss))
                Else
                    rsresults![youspitfire1l] = 0
            End If
            If (defspitfire2 + defspitfire2loss) > 0 Then
                    rsresults![youspitfire2l] = defspitfire2loss * (defSendArr(7, index) / (defspitfire2 + defspitfire2loss))
                Else
                    rsresults![youspitfire2l] = 0
            End If
            If (defspitfire5 + defspitfire5loss) > 0 Then
                    rsresults![youspitfire5l] = defspitfire5loss * (defSendArr(8, index) / (defspitfire5 + defspitfire5loss))
                Else
                    rsresults![youspitfire5l] = 0
            End If
            If (defhurricane + defhurricaneloss) > 0 Then
                    rsresults![youhurricanel] = defhurricaneloss * (defSendArr(9, index) / (defhurricane + defhurricaneloss))
                Else
                    rsresults![youhurricanel] = 0
            End If
            If (defLandingcraft + defLandingcraftloss) > 0 Then
                    rsresults![youlandingcraftl] = defLandingcraft * (defSendArr(13, index) / (defLandingcraft + defLandingcraftloss))
                Else
                    rsresults![youlandingcraftl] = 0
            End If
            If (defLandingship + defLandingshiploss) > 0 Then
                    rsresults![youlandingshipl] = defLandingshiploss * (defSendArr(14, index) / (defLandingship + defLandingshiploss))
                Else
                    rsresults![youlandingshipl] = 0
            End If
            If (deffrigates + deffrigatesloss) > 0 Then
                    rsresults![youfrigatesl] = deffrigatesloss * (defSendArr(15, index) / (deffrigates + deffrigatesloss))
                Else
                    rsresults![youfrigatesl] = 0
            End If
            If (defbattleships + defbattleshipsloss) > 0 Then
                    rsresults![youbattleshipl] = defbattleshipsloss * (defSendArr(16, index) / (defbattleships + defbattleshipsloss))
                Else
                    rsresults![youbattleshipl] = 0
            End If
            If (defcarrier + defcarrierloss) > 0 Then
                    rsresults![youcarrierl] = defcarrierloss * (defSendArr(17, index) / (defcarrier + defcarrierloss))
                Else
                    rsresults![youcarrierl] = 0
            End If
            If index = UBound(defUsernameArr) Then
                    rsresults![youaagunl] = defaagunsloss
                    rsresults![youpilboxesl] = pillboxeslost
                    rsresults![youturretl] = turretslost
                    rsresults![youballoonl] = defbarrageballoonsloss
                    rsresults![youminel] = MinesLost
                    rsresults![youaaguns] = defaagunsloss + defaagun
                    rsresults![youpilboxes] = pillboxeslost + defpillboxes
                    rsresults![youturrets] = turretslost + defturrets
                    rsresults![youballons] = defbarrageballoonsloss + defbarrageballoons
                    rsresults![youmines] = MinesLost + defwatermines
                Else
                    rsresults![youaagunl] = 0
                    rsresults![youpilboxesl] = 0
                    rsresults![youturretl] = 0
                    rsresults![youballoonl] = 0
                    rsresults![youminel] = 0
                    rsresults![youaaguns] = 0
                    rsresults![youpilboxes] = 0
                    rsresults![youturrets] = 0
                    rsresults![youballons] = 0
                    rsresults![youmines] = 0
            End If
            rsresults![youinfantrys] = defSendArr(0, index)
            rsresults![youcomandoess] = defSendArr(1, index)
            rsresults![youjeepss] = defSendArr(2, index)
            rsresults![youmatilda1s] = defSendArr(3, index)
            rsresults![youmatilda2s] = defSendArr(4, index)
            rsresults![youshermans] = defSendArr(5, index)
            rsresults![youspitfire1s] = defSendArr(6, index)
            rsresults![youspitfire2s] = defSendArr(7, index)
            rsresults![youspitfire5s] = defSendArr(8, index)
            rsresults![youhurricanes] = defSendArr(9, index)
            rsresults![youhalifaxs] = defSendArr(10, index)
            rsresults![youwellingtons] = defSendArr(11, index)
            rsresults![youlancasters] = defSendArr(12, index)
            rsresults![youlandingcrafts] = defSendArr(13, index)
            rsresults![youlandingships] = defSendArr(14, index)
            rsresults![youfrigatess] = defSendArr(15, index)
            rsresults![youbattleships] = defSendArr(16, index)
            rsresults![youcarriers] = defSendArr(17, index)
            
            rsresults![UserName] = defUsernameArr(index)
                
            rsresults![FarmsLost] = FarmsLost
            rsresults![MillsLost] = MillsLost
            rsresults![RefineriesLost] = RefineriesLost
            rsresults![MinesLost] = MinesLost
            rsresults![landlost] = totallost
            
            rsresults![attinfantrys] = attinfantry + attinfantryloss
            rsresults![attcomandoess] = attcommandoes + attcommandoesloss
            rsresults![attjeepss] = attjeeps + attjeepsloss
            rsresults![attmatilda1s] = attmatilda1 + attmatilda1loss
            rsresults![attmatilda2s] = attmatilda2 + attmatilda2loss
            rsresults![attshermans] = attsherman + attshermanloss
            rsresults![attspitfire1s] = attspitfire1 + attspitfire1loss
            rsresults![attspitfire2s] = attspitfire2 + attspitfire2loss
            rsresults![attspitfire5s] = attspitfire5 + attspitfire5loss
            rsresults![atthurricanes] = atthurricane + atthurricaneloss
            rsresults![attlancasters] = attlancaster + attlancasterloss
            rsresults![attwellingtons] = attwellington + attwellingtonloss
            rsresults![atthalifaxs] = atthalifax + atthalifaxloss
            rsresults![attfrigates] = attfrigates + attfrigatesloss
            rsresults![attbattleships] = attbattleships + attbattleshipsloss
            rsresults![attcarriers] = attcarrier + attcarrierloss
            
            rsresults![attinfantryl] = attinfantryloss
            rsresults![attcomandoesl] = attcommandoesloss
            rsresults![attjeepsl] = attjeepsloss
            rsresults![attmatilda1] = attmatilda1loss
            rsresults![attmatilda2] = attmatilda2loss
            rsresults![attshermanl] = attshermanloss
            rsresults![attspitfire1l] = attspitfire1loss
            rsresults![attspitfire2l] = attspitfire2loss
            rsresults![attspitfire5l] = attspitfire5loss
            rsresults![atthurricanel] = atthurricaneloss
            rsresults![attfrigatesl] = attfrigatesloss
            rsresults![attbattleshipl] = attbattleshipsloss
            rsresults![attcarrierl] = attcarrierloss
                        
            rsresults.Update
        Next
        rsresults.Close
        
        For index = 0 To intAttCounter
            main.SQLCon.RecordSource = "SELECT tblmovements.* FROM tblmovements WHERE orderNO = " & attIDArr(index)
            main.SQLCon.Refresh
            Set rsattack = main.SQLCon.Recordset
            If Not rsattack.EOF Then
                rsattack![status] = 3
                rsattack![eta] = rsattack![Ticktime]
                If attinfantry > 0 Then
                    rsattack![infantry] = attinfantry * (attSendArr(0, index) / (attinfantry + attinfantryloss))
                Else
                    rsattack![infantry] = 0
                End If
                If attcommandoes > 0 Then
                    rsattack![comandoes] = attcommandoes * (attSendArr(1, index) / (attcommandoes + attcommandoesloss))
                Else
                    rsattack![comandoes] = 0
                End If
                If attjeeps > 0 Then
                    rsattack![jeeps] = attjeeps * (attSendArr(2, index) / (attjeeps + attjeepsloss))
                Else
                    rsattack![jeeps] = 0
                End If
                If attmatilda1 > 0 Then
                    rsattack![matilda1] = attmatilda1 * (attSendArr(3, index) / (attmatilda1 + attmatilda1loss))
                Else
                    rsattack![matilda1] = 0
                End If
                If attmatilda2 > 0 Then
                    rsattack![matilda2] = attmatilda2 * (attSendArr(4, index) / (attmatilda2 + attmatilda2loss))
                Else
                    rsattack![matilda2] = 0
                End If
                If attsherman > 0 Then
                    rsattack![sherman] = attsherman * (attSendArr(5, index) / (attsherman + attshermanloss))
                Else
                    rsattack![sherman] = 0
                End If
                If attspitfire1 > 0 Then
                    rsattack![spitfire1] = attspitfire1 * (attSendArr(6, index) / (attspitfire1 + attspitfire1loss))
                Else
                    rsattack![spitfire1] = 0
                End If
                If attspitfire2 > 0 Then
                    rsattack![spitfire2] = attspitfire2 * (attSendArr(7, index) / (attspitfire2 + attspitfire2loss))
                Else
                    rsattack![spitfire2] = 0
                End If
                If attspitfire5 > 0 Then
                    rsattack![spitfire5] = attspitfire5 * (attSendArr(8, index) / (attspitfire5 + attspitfire5loss))
                Else
                    rsattack![spitfire5] = 0
                End If
                If atthurricane > 0 Then
                    rsattack![hurricane] = atthurricane * (attSendArr(9, index) / (atthurricane + atthurricaneloss))
                Else
                    rsattack![hurricane] = 0
                End If
                If attlancaster > 0 Then
                    rsattack![lancaster] = attlancaster * (attSendArr(12, index) / (attlancaster + attlancasterloss))
                Else
                    rsattack![lancaster] = 0
                End If
                If attwellington > 0 Then
                    rsattack![wellington] = attwellington * (attSendArr(11, index) / (attwellington + attwellingtonloss))
                Else
                    rsattack![wellington] = 0
                End If
                If atthalifax > 0 Then
                    rsattack![halifax] = atthalifax * (attSendArr(10, index) / (atthalifax + atthalifaxloss))
                Else
                    rsattack![halifax] = 0
                End If
                If attfrigates > 0 Then
                    rsattack![frigates] = attfrigates * (attSendArr(15, index) / (attfrigates + attfrigatesloss))
                Else
                    rsattack![frigates] = 0
                End If
                If attbattleships > 0 Then
                    rsattack![battleships] = attbattleships * (attSendArr(16, index) / (attbattleships + attbattleshipsloss))
                Else
                    rsattack![battleships] = 0
                End If
                If attcarrier > 0 Then
                    rsattack![carrier] = attcarrier * (attSendArr(17, index) / (attcarrier + attcarrierloss))
                Else
                    rsattack![carrier] = 0
                End If
                inttocontinent = rsattack![tocontinent]
                inttocountry = rsattack![tocountry]
                
                rsattack![tocontinent] = rsattack![fromcontinent]
                rsattack![tocountry] = rsattack![fromcountry]
                rsattack![fromcontinent] = inttocontinent
                rsattack![fromcountry] = inttocountry
                
                rsattack.Update
            End If

            Set rsattack = Nothing
        Next
        For index = 0 To intDefCounter
            If index = intDefCounter Then
                    Set rsdefend = main.SQLCon2.Recordset
                    
                    
                    If definfantry > 0 Then
                            rsdefend![numinfantry] = definfantry * (defSendArr(0, index) / (definfantry + definfantryloss))
                        Else
                            rsdefend![numinfantry] = 0
                    End If
                    If defcommandoes > 0 Then
                            rsdefend![numcommandos] = defcommandoes * (defSendArr(1, index) / (defcommandoes + defcommandoesloss))
                        Else
                            rsdefend![numcommandos] = 0
                    End If
                    If defjeeps > 0 Then
                            rsdefend![numjeeps] = defjeeps * (defSendArr(2, index) / (defjeeps + defjeepsloss))
                        Else
                            rsdefend![numjeeps] = 0
                    End If
                    If defmatilda1 > 0 Then
                            rsdefend![nummat1] = defmatilda1 * (defSendArr(3, index) / (defmatilda1 + defmatilda1loss))
                        Else
                            rsdefend![nummat1] = 0
                    End If
                    If defmatilda2 > 0 Then
                            rsdefend![nummat2] = defmatilda2 * (defSendArr(4, index) / (defmatilda2 + defmatilda2loss))
                        Else
                            rsdefend![nummat2] = 0
                    End If
                    If defsherman > 0 Then
                            rsdefend![numsherman] = defsherman * (defSendArr(5, index) / (defsherman + defshermanloss))
                        Else
                            rsdefend![numsherman] = 0
                    End If
                    If defspitfire1 > 0 Then
                            rsdefend![numspit1] = defspitfire1 * (defSendArr(6, index) / (defspitfire1 + defspitfire1loss))
                        Else
                            rsdefend![numspit1] = 0
                    End If
                    If defspitfire2 > 0 Then
                            rsdefend![numspit2] = defspitfire2 * (defSendArr(7, index) / (defspitfire2 + defspitfire2loss))
                        Else
                            rsdefend![numspit2] = 0
                    End If
                    If defspitfire5 > 0 Then
                            rsdefend![numspit5] = defspitfire5 * (defSendArr(8, index) / (defspitfire5 + defspitfire5loss))
                        Else
                            rsdefend![numspit5] = 0
                    End If
                    If defhurricane > 0 Then
                            rsdefend![numhurr] = defhurricane * (defSendArr(9, index) / (defhurricane + defhurricaneloss))
                        Else
                            rsdefend![numhurr] = 0
                    End If
                    If defLandingcraft > 0 Then
                            rsdefend![numlandingcraft] = defLandingcraft * (defSendArr(13, index) / (defLandingcraft + defLandingcraftloss))
                        Else
                            rsdefend![numlandingcraft] = 0
                    End If
                    If defLandingship > 0 Then
                            rsdefend![numlandingship] = defLandingship * (defSendArr(14, index) / (defLandingship + defLandingshiploss))
                        Else
                            rsdefend![numlandingship] = 0
                    End If
                    If deffrigates > 0 Then
                            rsdefend![numfrigates] = deffrigates * (defSendArr(15, index) / (deffrigates + deffrigatesloss))
                        Else
                            rsdefend![numfrigates] = 0
                    End If
                    If defbattleships > 0 Then
                            rsdefend![numbattleship] = defbattleships * (defSendArr(16, index) / (defbattleships + defbattleshipsloss))
                        Else
                            rsdefend![numbattleship] = 0
                    End If
                    If defcarrier > 0 Then
                            rsdefend![numcarrier] = defcarrier * (defSendArr(17, index) / (defcarrier + defcarrierloss))
                        Else
                            rsdefend![numcarrier] = 0
                    End If
                    rsdefend![numaaguns] = defaagun
                    rsdefend![numpilboxes] = defpillboxes
                    rsdefend![numturrets] = defturrets
                    rsdefend![numballoons] = defbarrageballoons
                    rsdefend![numwatermines] = defwatermines
                    
                    
                    rsdefend![amountland] = rsdefend![amountland] - totallost
                    
                    rsdefend![nummines] = rsdefend![nummines] - MinesLost
                    rsdefend![nummills] = rsdefend![nummills] - MillsLost
                    rsdefend![numfarms] = rsdefend![numfarms] - FarmsLost
                    rsdefend![numrefineries] = rsdefend![numrefineries] - RefineriesLost
                    
                    rsdefend.Update
            Else
            main.SQLCon.RecordSource = "SELECT tblmovements.* FROM tblmovements WHERE orderNO = " & defIDArr(index)
            main.SQLCon.Refresh
            Set rsdefend = main.SQLCon.Recordset
            If Not rsdefend.EOF Then
                rsdefend![status] = 3
                rsdefend![eta] = rsdefend![Ticktime]
                If definfantry > 0 Then
                    rsdefend![infantry] = definfantry * (defSendArr(0, index) / (definfantry + definfantryloss))
                Else
                    rsdefend![infantry] = 0
                End If
                If defcommandoes > 0 Then
                    rsdefend![comandoes] = defcommandoes * (defSendArr(1, index) / (defcommandoes + defcommandoesloss))
                Else
                    rsdefend![comandoes] = 0
                End If
                If defjeeps > 0 Then
                    rsdefend![jeeps] = defjeeps * (defSendArr(2, index) / (defjeeps + defjeepsloss))
                Else
                    rsdefend![jeeps] = 0
                End If
                If defmatilda1 > 0 Then
                    rsdefend![matilda1] = defmatilda1 * (defSendArr(3, index) / (defmatilda1 + defmatilda1loss))
                Else
                    rsdefend![matilda1] = 0
                End If
                If defmatilda2 > 0 Then
                    rsdefend![matilda2] = defmatilda2 * (defSendArr(4, index) / (defmatilda2 + defmatilda2loss))
                Else
                    rsdefend![matilda2] = 0
                End If
                If defsherman > 0 Then
                    rsdefend![sherman] = defsherman * (defSendArr(5, index) / (defsherman + defshermanloss))
                Else
                    rsdefend![sherman] = 0
                End If
                If defspitfire1 > 0 Then
                    rsdefend![spitfire1] = defspitfire1 * (defSendArr(6, index) / (defspitfire1 + defspitfire1loss))
                Else
                    rsdefend![spitfire1] = 0
                End If
                If defspitfire2 > 0 Then
                    rsdefend![spitfire2] = defspitfire2 * (defSendArr(7, index) / (defspitfire2 + defspitfire2loss))
                Else
                    rsdefend![spitfire2] = 0
                End If
                If defspitfire5 > 0 Then
                    rsdefend![spitfire5] = defspitfire5 * (defSendArr(8, index) / (defspitfire5 + defspitfire5loss))
                Else
                    rsdefend![spitfire5] = 0
                End If
                If defhurricane > 0 Then
                    rsdefend![hurricane] = defhurricane * (defSendArr(9, index) / (defhurricane + defhurricaneloss))
                Else
                    rsdefend![hurricane] = 0
                End If
                If deffrigates > 0 Then
                    rsdefend![frigates] = deffrigates * (defSendArr(15, index) / (deffrigates + deffrigatesloss))
                Else
                    rsdefend![frigates] = 0
                End If
                If defbattleships > 0 Then
                    rsdefend![battleships] = defbattleships * (defSendArr(16, index) / (defbattleships + defbattleshipsloss))
                Else
                    rsdefend![battleships] = 0
                End If
                If defcarrier > 0 Then
                    rsdefend![carrier] = defcarrier * (defSendArr(17, index) / (defcarrier + defcarrierloss))
                Else
                    rsdefend![carrier] = 0
                End If
                inttocontinent = rsdefend![tocontinent]
                inttocountry = rsdefend![tocountry]
                
                rsdefend![tocontinent] = rsdefend![fromcontinent]
                rsdefend![tocountry] = rsdefend![fromcountry]
                rsdefend![fromcontinent] = inttocontinent
                rsdefend![fromcountry] = inttocountry
            
                rsdefend.Update
            End If
            End If
        Next
        
        Do While totallost > 0
            Randomize
            num = CInt(Rnd * intAttCounter)
            landGainArr(num) = landGainArr(num) + 1
            totallost = totallost - 1
        Loop
        
        For index = 0 To intAttCounter
            If landGainArr(index) > 0 Then
                    main.SQLCon.RecordSource = "SELECT amountland FROM tblaccounts WHERE username = '" & attUsernameArr(index) & "';"
                    
                    main.SQLCon.Refresh
                    
                    If Not main.SQLCon.Recordset.EOF Then
                            main.SQLCon.Recordset![amountland] = main.SQLCon.Recordset![amountland] + landGainArr(index)
                            
                            main.SQLCon.Recordset.Update
                    End If
                    
                    main.SQLCon.Recordset.Close
            End If
        Next
        main.SQLCon.RecordSource = "SELECT * FROM tblnews ORDER BY newsid DESC;"
            
        main.SQLCon.Refresh
        If main.SQLCon.Recordset.EOF Then
                resultno = 1
            Else
                resultno = main.SQLCon.Recordset![newsid] + 1
        End If
        For index = 0 To intAttCounter
            main.SQLCon.Recordset.AddNew
            main.SQLCon.Recordset![newsid] = resultno
            main.SQLCon.Recordset![UserName] = attUsernameArr(index)
            main.SQLCon.Recordset![timemade] = Now()
            main.SQLCon.Recordset![message] = "Your military forces have engaged in offencive combat and have left the following <A HREF=view_attack.asp?id=" & attReportIDArr(index) & ">combat report</A> for you.  "
            main.SQLCon.Recordset![unread] = 1
            main.SQLCon.Recordset.Update
            resultno = resultno + 1
        Next
            
        For index = 0 To intDefCounter
            main.SQLCon.Recordset.AddNew
            main.SQLCon.Recordset![newsid] = resultno
            main.SQLCon.Recordset![UserName] = defUsernameArr(index)
            main.SQLCon.Recordset![timemade] = Now()
            main.SQLCon.Recordset![message] = "Your military forces have been engaged in defenvice combat and have left the following <A HREF=view_attack.asp?id=" & defReportIDArr(index) & ">combat report</A> for you.  "
            main.SQLCon.Recordset![unread] = 1
            main.SQLCon.Recordset.Update
            resultno = resultno + 1
        Next
        
        main.SQLCon.Recordset.Close
    
    
        End If
    main.SQLCon2.Recordset.MoveNext
    Loop
    
    main.SQLCon.RecordSource = "SELECT tblmovements.* FROM tblmovements WHERE ETA <= 0 AND Status > 1"
    main.SQLCon.Refresh
    
    Set rsreturn = main.SQLCon.Recordset
    If Not rsreturn.EOF Then
        Do While Not rsreturn.EOF
            If rsreturn![status] = 2 Then
                    rsreturn![eta] = rsreturn![Ticktime]
                    rsreturn![status] = 3
                    inttocontinent = rsreturn![fromcontinent]
                    inttocountry = rsreturn![fromcountry]
                    rsreturn![fromcontinent] = rsreturn![tocontinent]
                    rsreturn![fromcountry] = rsreturn![tocountry]
                    rsreturn![tocontinent] = inttocontinent
                    rsreturn![tocountry] = inttocountry
                    
                    rsreturn.Update
                Else
                    main.SQLCon2.RecordSource = "SELECT numinfantry, numcommandos, numjeeps, numsherman, nummat1, nummat2, numspit1, numspit2, numspit5, numhurr, numfrigates, numbattleship, numcarrier, numlancaster, numwellington, numhalifax, numlandingship, numlandingcraft FROM tblaccounts WHERE username = '" & rsreturn![UserName] & "';"
                    main.SQLCon2.Refresh
                    Set rsaccount = main.SQLCon2.Recordset
                    If Not rsaccount.EOF Then
                        rsaccount![numinfantry] = rsaccount![numinfantry] + rsreturn![infantry]
                        rsaccount![numcommandos] = rsaccount![numcommandos] + rsreturn![comandoes]
                        rsaccount![numjeeps] = rsaccount![numjeeps] + rsreturn![jeeps]
                        rsaccount![nummat1] = rsaccount![nummat1] + rsreturn![matilda1]
                        rsaccount![nummat2] = rsaccount![nummat2] + rsreturn![matilda2]
                        rsaccount![numsherman] = rsaccount![numsherman] + rsreturn![sherman]
                        rsaccount![numspit1] = rsaccount![numspit1] + rsreturn![spitfire1]
                        rsaccount![numspit2] = rsaccount![numspit2] + rsreturn![spitfire2]
                        rsaccount![numspit5] = rsaccount![numspit5] + rsreturn![spitfire5]
                        rsaccount![numhurr] = rsaccount![numhurr] + rsreturn![hurricane]
                        rsaccount![numfrigates] = rsaccount![numfrigates] + rsreturn![frigates]
                        rsaccount![numbattleship] = rsaccount![numbattleship] + rsreturn![battleships]
                        rsaccount![numcarrier] = rsaccount![numcarrier] + rsreturn![carrier]
                        rsaccount![numlancaster] = rsaccount![numlancaster] + rsreturn![lancaster]
                        rsaccount![numwellington] = rsaccount![numwellington] + rsreturn![wellington]
                        rsaccount![numhalifax] = rsaccount![numhalifax] + rsreturn![halifax]
                        rsaccount![numlandingship] = rsaccount![numlandingship] + rsreturn![landingship]
                        rsaccount![numlandingcraft] = rsaccount![numlandingcraft] + rsreturn![landingcraft]
                        rsaccount.Update
                    End If
                    rsreturn.Delete
            End If
        rsreturn.MoveNext
        Loop
    End If

End Sub

