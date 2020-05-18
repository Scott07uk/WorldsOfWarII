Attribute VB_Name = "Attackscript"
Option Explicit

Sub attack()
    Dim resultno As Integer
    Dim attloss As Long
    Dim defloss As Long
    Dim MinesLost As Integer
    Dim FarmsLost As Integer
    Dim MillsLost As Integer
    Dim RefineriesLost As Integer
    Dim ratio As Double
    Dim rsattack
    Dim attacknumber As Integer
    Dim rsdefend
    
    Dim attspitfire1 As Integer
    Dim attspitfire2 As Integer
    Dim attspitfire5 As Integer
    Dim atthurricane As Integer
    Dim attlancaster As Integer
    Dim attwellington As Integer
    Dim atthalifax As Integer
    Dim attspitfire1loss As Integer
    Dim attspitfire2loss As Integer
    Dim attspitfire5loss As Integer
    Dim atthurricaneloss As Integer
    Dim attlancasterloss As Integer
    Dim attwellingtonloss As Integer
    Dim atthalifaxloss As Integer
    Dim totallost As Integer
    
    Dim attunittotal As Long
    Dim attforcetotal As Long
    Dim defunittotal As Long
    Dim defforcetotal As Long
    
    Dim defspitfire1 As Integer
    Dim defspitfire2 As Integer
    Dim defspitfire5 As Integer
    Dim defhurricane As Integer
    Dim defaagun As Integer
    Dim defturrets As Integer
    Dim defpillboxes As Integer
    Dim attbombingtotal As Integer
    Dim turretslost As Integer
    Dim pillboxeslost As Integer
    Dim num
    
    Dim defspitfire1loss As Integer
    Dim defspitfire2loss As Integer
    Dim defspitfire5loss As Integer
    Dim defhurricaneloss As Integer
    Dim defaagunsloss As Integer
    
    Dim amountdefenceland As Integer
    Dim numshiplost As Integer
    Dim nummineslost As Integer
    
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
    Dim attbarrageballoons As Long
    
    Dim definfantryloss As Integer
    Dim defcommandoesloss As Integer
    Dim defjeepsloss As Integer
    Dim defmatilda1loss As Integer
    Dim defmatilda2loss As Integer
    Dim defshermanloss As Integer
    Dim defLandingcraftloss As Integer
    Dim defLandingshiploss As Integer
    Dim deffrigatesloss As Integer
    Dim defbattleshipsloss As Integer
    Dim defcarrierloss As Integer
    Dim defbarrageballoonsloss As Integer
    
    Dim attinfantryloss As Integer
    Dim attcommandoesloss As Integer
    Dim attjeepsloss As Integer
    Dim attmatilda1loss As Integer
    Dim attmatilda2loss As Integer
    Dim attshermanloss As Integer
    Dim attLandingcraftloss As Integer
    Dim attLandingshiploss As Integer
    Dim attfrigatesloss As Integer
    Dim attbattleshipsloss As Integer
    Dim attcarrierloss As Integer
    Dim rsresults
    Dim rsreturn
    Dim rsaccount
    
    main.SQLCon.RecordSource = "SELECT tblmovements.ETA FROM tblmovements;"
    main.SQLCon.Refresh
    
    Set rsattack = main.SQLCon.Recordset
    If Not rsattack.EOF Then
        Do While Not rsattack.EOF
            rsattack![ETA] = rsattack![ETA] - 1
            rsattack.Update
            rsattack.MoveNext
        Loop
    End If
    
    main.SQLCon.RecordSource = "SELECT tblmovements.* FROM tblmovements WHERE ETA <= 0 AND Status = 1"
    main.SQLCon.Refresh
    
    Set rsattack = main.SQLCon.Recordset
    If Not rsattack.EOF Then
        Do While Not rsattack.EOF
                attacknumber = attacknumber + 1
            
                Print #1, "+----------------------------------------+" & vbCrLf
                Print #1, "| ATTACK DATA " & attacknumber & " |" & vbCrLf
                Print #1, "+----------------------------------------+" & vbCrLf & vbCrLf
    
                attspitfire1 = rsattack![spitfire1]
                attspitfire2 = rsattack![spitfire2]
                attspitfire5 = rsattack![spitfire5]
                atthurricane = rsattack![hurricane]
                attlancaster = rsattack![lancaster]
                attwellington = rsattack![wellington]
                atthalifax = rsattack![halifax]
                
                attunittotal = attspitfire1 + attspitfire2 + attspitfire5 + atthurricane
                
                attspitfire2force = rsattack![spitfire2] * 2
                attspitfire5force = rsattack![spitfire5] * 3
                atthurricaneforce = rsattack![hurricane] * 2
                
                attforcetotal = attspitfire1 + attspitfire2force + attspitfire5force + atthurricaneforce
                
                main.SQLCon2.RecordSource = "SELECT tblaccounts.* FROM tblaccounts WHERE continent = " & rsattack![ToContinent] & "AND countryID = " & rsattack![ToCountry]
                main.SQLCon2.Refresh
                Set rsdefend = main.SQLCon2.Recordset
                If Not rsdefend.EOF Then
                
                    'PLANES VS PLANES
                    defspitfire1 = rsdefend![numspit1]
                    defspitfire2 = rsdefend![numspit2]
                    defspitfire5 = rsdefend![numspit5]
                    defhurricane = rsdefend![numhurr]
                    
                    defunittotal = defspitfire1 + defspitfire2 + defspitfire5 + defhurricane
                    
                    defspitfire2force = rsdefend![numspit2] * 2
                    defspitfire5force = rsdefend![numspit5] * 3
                    defhurricaneforce = rsdefend![numhurr] * 2
                    
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
                        defaagun = rsdefend![numaaguns]
                        
                        If defaagun > 0 Then
                            defunittotal = defaagun
                            defforcetotal = defaagun
                            
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
                                
                            Print #1, "ratio bomber " & ratio
                            Print #1, "deftotal bomber " & defunittotal
                            Print #1, "atttotal bomber " & attunittotal
                            Print #1, "attloss bomber " & attloss
                            Print #1, "defloss bomber " & defloss
                        End If
                        
                        'PLANES VS STUFF THAT CANT DEFEND ITS SELF
                        defturrets = rsdefend![numturrets]
                        defpillboxes = rsdefend![numpilboxes]
                        attbombingtotal = attlancaster + atthalifax + attwellington
                        attbombingtotal = attbombingtotal / 4
                        FarmsLost = rsdefend![numfarms] * (attbombingtotal / 10)
                        MillsLost = rsdefend![nummills] * (attbombingtotal / 12)
                        MinesLost = rsdefend![nummines] * (attbombingtotal / 14)
                        RefineriesLost = rsdefend![numrefineries] * (attbombingtotal / 12)
                        
                        If FarmsLost > rsdefend![numfarms] Then
                            FarmsLost = 0
                        End If
                        If MillsLost > rsdefend![nummills] Then
                            MillsLost = 0
                        End If
                        If MinesLost > rsdefend![nummines] Then
                            MinesLost = 0
                        End If
                        If RefineriesLost > rsdefend![numrefineries] Then
                            RefineriesLost = 0
                        End If
                        
                        If attbombingtotal > 0 Then
                            If defturrets > 0 Then
                                turretslost = defturrets / (attbombingtotal / 2)
                            Else
                                turretslost = 0
                            End If
                            If defpillboxes > 0 And attbombingtotal > 0 Then
                                pillboxeslost = defpillboxes / (attbombingtotal / 2)
                            Else
                                pillboxeslost = 0
                            End If
                        End If
                        totallost = MinesLost + FarmsLost + MillsLost + RefineriesLost
                        Randomize
                        num = (Rnd * 0.2) + 0.4
                        totallost = totallost * num
                        
                        
                        Print #1, rsdefend![nummills]
                        
                        Print #1, "farms lost" & FarmsLost
                        Print #1, "Mills lost" & MillsLost
                        Print #1, "mines lost" & MinesLost
                        Print #1, "refineries lost" & RefineriesLost
                        Print #1, "Turrets lost" & turretslost
                        Print #1, "Pillboxes lost" & pillboxeslost
                    End If
                    'RUBBISH BOATS VS MINES THAT DONT DO ANYTHING
                    If rsdefend![numwatermines] > 0 And (rsattack![landingship] + rsattack![landingcraft]) > 0 Then
                        amountdefenceland = rsdefend![AmountLand]
                        amountdefenceland = Sqr(amountdefenceland)
                        amountdefenceland = amountdefenceland * 3.141592654
                        amountdefenceland = amountdefenceland / rsdefend![numwatermines]
                        amountdefenceland = amountdefenceland / 4
                        numshiplost = amountdefenceland * (rsattack![landingship] + rsattack![landingcraft])
                        nummineslost = amountdefenceland * (rsattack![landingship] + rsattack![landingcraft])
                    End If
                    
                    Print #1, "Ship lost " & numshiplost
                    Print #1, "Mines lost " & nummineslost
                    
                    'DUDES VS DUDES
                    definfantry = rsdefend![numinfantry]
                    defcommandoes = rsdefend![numcommandos]
                    defjeeps = rsdefend![numjeeps]
                    defmatilda1 = rsdefend![nummat1]
                    defmatilda2 = rsdefend![nummat2]
                    defsherman = rsdefend![numsherman]
                    defLandingcraft = rsdefend![numlandingcraft]
                    defLandingship = rsdefend![numlandingship]
                    deffrigates = rsdefend![numfrigates]
                    defbattleships = rsdefend![numbattleship]
                    defcarrier = rsdefend![numcarrier]
                    defturrets = rsdefend![numturrets]
                    defpillboxes = rsdefend![numpilboxes]
                    defbarrageballoons = rsdefend![numballoons]
                    
                    defunittotal = definfantry + defcommandoes + defjeeps + defmatilda1 + defmatilda2 + defsherman + deffrigates + defbattleships + defcarrier + defturrets + defpillboxes + defbarrageballoons
                
                    defcommandoesforce = defcommandoes * 2
                    defmatilda1force = defmatilda1 * 2
                    defmatilda2force = defmatilda2 * 3
                    defshermanforce = defsherman * 2
                    defbattleshipsforce = defbattleships * 2
                    defcarrierforce = defcarrier * 2
                    
                    defforcetotal = definfantry + defcommandoesforce + defjeeps + defmatilda1force + defmatilda2force + defshermanforce + deffrigates + defbattleshipsforce + defcarrierforce + defturrets + defpillboxes + defbarrageballoons
                    
                    attinfantry = rsattack![infantry]
                    attcommandoes = rsattack![comandoes]
                    attjeeps = rsattack![jeeps]
                    attmatilda1 = rsattack![matilda1]
                    attmatilda2 = rsattack![matilda2]
                    attsherman = rsattack![sherman]
                    attfrigates = rsattack![frigates]
                    attbattleships = rsattack![battleships]
                    attcarrier = rsattack![carrier]
                    
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
                    If defforcetotal > 0 And attforcetotal > 0 Then
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
                    
                    main.SQLCon.RecordSource = "SELECT tblattackResults.* FROM tblattackresults ORDER BY reportID DESC;"
                    main.SQLCon.Refresh
                    
                    Set rsresults = main.SQLCon.Recordset
                        resultno = rsresults![reportID] + 1
                        rsresults.AddNew
                        rsresults![reportID] = resultno
                        rsresults![countryID] = rsdefend![countryID]
                        rsresults![ContinetID] = rsdefend![continent]
                            
                        rsresults![definfantrys] = rsdefend![numinfantry]
                        rsresults![defcomandoess] = rsdefend![numcommandos]
                        rsresults![defjeepss] = rsdefend![numjeeps]
                        rsresults![defmatilda1s] = rsdefend![nummat1]
                        rsresults![defmatilda2s] = rsdefend![nummat2]
                        rsresults![defshermans] = rsdefend![numsherman]
                        rsresults![defspitfire1s] = rsdefend![numspit1]
                        rsresults![defspitfire2s] = rsdefend![numspit2]
                        rsresults![defspitfire5s] = rsdefend![numspit5]
                        rsresults![defhurricanes] = rsdefend![numhurr]
                        rsresults![deffrigatess] = rsdefend![numfrigates]
                        rsresults![defbattleships] = rsdefend![numbattleship]
                        rsresults![defcarriers] = rsdefend![numcarrier]
                        rsresults![defaagunss] = rsdefend![numaaguns]
                        rsresults![defpilboxess] = rsdefend![numpilboxes]
                        rsresults![defturretss] = rsdefend![numturrets]
                        rsresults![defballoonss] = rsdefend![numballoons]
                        rsresults![defminess] = rsdefend![numwatermines]
                        
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
                        
                        rsresults![youinfantryl] = definfantryloss
                        rsresults![youcomandoesl] = defcommandoesloss
                        rsresults![youjeepsl] = defjeepsloss
                        rsresults![youmatilda1l] = defmatilda1loss
                        rsresults![youmatilda2l] = defmatilda2loss
                        rsresults![youshermanl] = defshermanloss
                        rsresults![youspitfire1l] = defspitfire1loss
                        rsresults![youspitfire2l] = defspitfire2loss
                        rsresults![youspitfire5l] = defspitfire5loss
                        rsresults![youhurricanel] = defhurricaneloss
                        rsresults![youfrigatesl] = deffrigatesloss
                        rsresults![youbattleshipl] = defbattleshipsloss
                        rsresults![youcarrierl] = defcarrierloss
                        rsresults![youaagunl] = defaagunsloss
                        rsresults![youpilboxesl] = pillboxeslost
                        rsresults![youturretl] = turretslost
                        rsresults![youballoonl] = defbarrageballoonsloss
                        rsresults![youminel] = nummineslost
                        
                        rsresults![youinfantrys] = rsdefend![numinfantry]
                        rsresults![youcomandoess] = rsdefend![numcommandos]
                        rsresults![youjeepss] = rsdefend![numjeeps]
                        rsresults![youmatilda1s] = rsdefend![nummat1]
                        rsresults![youmatilda2s] = rsdefend![nummat2]
                        rsresults![youshermans] = rsdefend![numsherman]
                        rsresults![youspitfire1s] = rsdefend![numspit1]
                        rsresults![youspitfire2s] = rsdefend![numspit2]
                        rsresults![youspitfire5s] = rsdefend![numspit5]
                        rsresults![youhurricanes] = rsdefend![numhurr]
                        rsresults![youfrigatess] = rsdefend![numfrigates]
                        rsresults![youbattleships] = rsdefend![numbattleship]
                        rsresults![youcarriers] = rsdefend![numcarrier]
                        rsresults![youaaguns] = rsdefend![numaaguns]
                        rsresults![youpilboxes] = rsdefend![numpilboxes]
                        rsresults![youturrets] = rsdefend![numturrets]
                        rsresults![youballons] = rsdefend![numballoons]
                        rsresults![youmines] = rsdefend![numwatermines]
                        rsresults![UserName] = rsdefend![UserName]
                            
                        rsresults![FarmsLost] = FarmsLost
                        rsresults![MillsLost] = MillsLost
                        rsresults![RefineriesLost] = RefineriesLost
                        rsresults![MinesLost] = MinesLost
                        rsresults![landlost] = totallost
                        
                        rsresults![attinfantrys] = rsattack![infantry]
                        rsresults![attcomandoess] = rsattack![comandoes]
                        rsresults![attjeepss] = rsattack![jeeps]
                        rsresults![attmatilda1s] = rsattack![matilda1]
                        rsresults![attmatilda2s] = rsattack![matilda2]
                        rsresults![attshermans] = rsattack![sherman]
                        rsresults![attspitfire1s] = rsattack![spitfire1]
                        rsresults![attspitfire2s] = rsattack![spitfire2]
                        rsresults![attspitfire5s] = rsattack![spitfire5]
                        rsresults![atthurricanes] = rsattack![hurricane]
                        rsresults![attlancasters] = rsattack![lancaster]
                        rsresults![attwellingtons] = rsattack![wellington]
                        rsresults![atthalifaxs] = rsattack![halifax]
                        rsresults![attfrigates] = rsattack![frigates]
                        rsresults![attbattleships] = rsattack![battleships]
                        rsresults![attcarriers] = rsattack![carrier]
                        
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
                    rsresults.Close
                    
                    rsattack![status] = 3
                    rsattack![ETA] = rsattack![Ticktime]
                    rsattack![infantry] = rsattack![infantry] - attinfantryloss
                    rsattack![comandoes] = rsattack![comandoes] - attcommandoesloss
                    rsattack![jeeps] = rsattack![jeeps] - attjeepsloss
                    rsattack![matilda1] = rsattack![matilda1] - attmatilda1loss
                    rsattack![matilda2] = rsattack![matilda2] - attmatilda2loss
                    rsattack![sherman] = rsattack![sherman] - attshermanloss
                    rsattack![spitfire1] = rsattack![spitfire1] - attspitfire1loss
                    rsattack![spitfire2] = rsattack![spitfire2] - attspitfire2loss
                    rsattack![spitfire5] = rsattack![spitfire5] - attspitfire5loss
                    rsattack![hurricane] = rsattack![hurricane] - atthurricaneloss
                    rsattack![lancaster] = rsattack![lancaster] - attlancasterloss
                    rsattack![wellington] = rsattack![wellington] - attwellingtonloss
                    rsattack![halifax] = rsattack![halifax] - atthalifaxloss
                    rsattack![frigates] = rsattack![frigates] - attfrigatesloss
                    rsattack![battleships] = rsattack![battleships] - attbattleshipsloss
                    rsattack![carrier] = rsattack![carrier] - attcarrierloss
                    rsattack.Update
                    
                    rsdefend![numinfantry] = definfantry
                    rsdefend![numcommandos] = defcommandoes
                    rsdefend![numjeeps] = defjeeps
                    rsdefend![nummat1] = defmatilda1
                    rsdefend![nummat2] = defmatilda2
                    rsdefend![numsherman] = defsherman
                    rsdefend![numspit1] = defspitfire1
                    rsdefend![numspit2] = defspitfire2
                    rsdefend![numspit5] = defspitfire5
                    rsdefend![numhurr] = defhurricane
                    rsdefend![numfrigates] = deffrigates
                    rsdefend![numbattleship] = defbattleships
                    rsdefend![numcarrier] = defcarrierloss
                    rsdefend![numfarms] = rsdefend![numfarms] - FarmsLost
                    rsdefend![nummills] = rsdefend![nummills] - MillsLost
                    rsdefend![numrefineries] = rsdefend![numrefineries] - RefineriesLost
                    rsdefend![nummines] = rsdefend![nummines] - MinesLost
                    rsdefend![numpilboxes] = rsdefend![numpilboxes] - pillboxeslost
                    rsdefend![numturrets] = rsdefend![numturrets] - turretslost
                    rsdefend![numwatermines] = rsdefend![numwatermines] - nummineslost
                    rsdefend![numaaguns] = rsdefend![numaaguns] - defaagunsloss
                    rsdefend![numballoons] = rsdefend![numballoons] - defbarrageballoonsloss
                    rsdefend.Update
                End If
        rsattack.MoveNext
        Loop
    End If
    
    main.SQLCon.RecordSource = "SELECT tblmovements.* FROM tblmovements WHERE ETA <= 0 AND Status = 3"
    main.SQLCon.Refresh
    
    Set rsreturn = main.SQLCon.Recordset
    If Not rsreturn.EOF Then
        Do While Not rsreturn.EOF
            main.SQLCon2.RecordSource = "SELECT tblaccounts.* FROM tblaccounts WHERE username = '" & rsreturn![UserName] & "';"
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
        rsreturn.MoveNext
        Loop
    End If
    rsreturn.Close
    rsattack.Close
End Sub

