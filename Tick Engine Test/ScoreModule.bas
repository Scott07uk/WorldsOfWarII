Attribute VB_Name = "ScoreModule"
Public Sub calScore()
Dim dblWorkingScore As Double
Dim lngFinalScore As Long
Dim lngInfantry As Long
Dim lngCommandoes As Long
Dim lngJeeps As Long
Dim lngMatilda1 As Long
Dim lngMatilda2 As Long
Dim lngSherman As Long
Dim lngSpitfire1 As Long
Dim lngSpitfire2 As Long
Dim lngSpitfire5 As Long
Dim lngHurricane As Long
Dim lngHalifax As Long
Dim lngWellington As Long
Dim lngLancaster As Long
Dim lngLandingCraft As Long
Dim lngLandingShip As Long
Dim lngFrigate As Long
Dim lngBattleShip As Long
main.SQLCon.RecordSource = "SELECT username, gold, numinfantry, numcommandos, numjeeps, nummat1, nummat2, numsherman, numspit1, numspit2, numspit5, numhurr, numlancaster, numwellington, numhalifax, numlandingcraft, numlandingship, numfrigates, numbattleship, numcarrier, numaaguns, numpilboxes, numturrets, numballoons, numwatermines, amountland, nummp, numspies, numexplorers, numfarms, nummines, numrefineries, nummills, score FROM tblaccounts;"

main.SQLCon.Refresh

Do While Not main.SQLCon.Recordset.EOF
    lngInfantry = 0
    lngCommandoes = 0
    lngJeeps = 0
    lngMatilda1 = 0
    lngMatilda2 = 0
    lngSherman = 0
    lngSpitfire1 = 0
    lngSpitfire2 = 0
    lngSpitfire5 = 0
    lngHurricane = 0
    lngHalifax = 0
    lngWellington = 0
    lngLancaster = 0
    lngLandingCraft = 0
    lngLandingShip = 0
    lngFrigate = 0
    lngBattleShip = 0
    main.SQLCon2.RecordSource = "SELECT infantry, comandoes, jeeps, matilda1, matilda2, sherman, spitfire1, spitfire2, spitfire5, hurricane, lancaster, wellington, halifax, landingcraft, landingship, frigates, battleships, carrier FROM tblMovements WHERE username = '" & main.SQLCon.Recordset![UserName] & "';"
    
    main.SQLCon2.Refresh
    
    Do While Not main.SQLCon2.Recordset.EOF
        lngInfantry = lngInfantry + CLng(main.SQLCon2.Recordset![infantry])
        lngCommandoes = lngCommandoes + CLng(main.SQLCon2.Recordset![comandoes])
        lngJeeps = lngJeeps + CLng(main.SQLCon2.Recordset![jeeps])
        lngMatilda1 = lngMatilda1 + CLng(main.SQLCon2.Recordset![matilda1])
        lngMatilda2 = lngMatilda2 + CLng(main.SQLCon2.Recordset![matilda2])
        lngSherman = lngSherman + CLng(main.SQLCon2.Recordset![sherman])
        lngSpitfire1 = lngSpitfire1 + CLng(main.SQLCon2.Recordset![spitfire1])
        lngSpitfire2 = lngSpitfire2 + CLng(main.SQLCon2.Recordset![spitfire2])
        lngSpitfire5 = lngSpitfire5 + CLng(main.SQLCon2.Recordset![spitfire5])
        lngHurricane = lngHurricane + CLng(main.SQLCon2.Recordset![hurricane])
        lngHalifax = lngHalifax + CLng(main.SQLCon2.Recordset![halifax])
        lngWellington = lngWellington + CLng(main.SQLCon2.Recordset![wellington])
        lngLancaster = lngLancaster + CLng(main.SQLCon2.Recordset![lancaster])
        lngLandingCraft = lngLandingCraft + CLng(main.SQLCon2.Recordset![landingcraft])
        lngLandingShip = lngLandingShip + CLng(main.SQLCon2.Recordset![landingship])
        lngFrigate = lngFrigate + CLng(main.SQLCon2.Recordset![frigates])
        lngBattleships = lngBattleships + CLng(main.SQLCon2.Recordset![battleships])
        
        main.SQLCon2.Recordset.MoveNext
    Loop
    main.SQLCon2.Recordset.Close
    
    dblWorkingScore = CLng(main.SQLCon.Recordset![gold]) + CLng(main.SQLCon.Recordset![amountland])
    
    dblWorkingScore = dblWorkingScore + CDbl((CLng(main.SQLCon.Recordset![numinfantry]) + lngInfantry) / 1000)
    dblWorkingScore = dblWorkingScore + CDbl((CLng(main.SQLCon.Recordset![numcommandos]) + lngCommandoes) / 500)
    dblWorkingScore = dblWorkingScore + CDbl((CLng(main.SQLCon.Recordset![numjeeps]) + lngJeeps) / 450)
    dblWorkingScore = dblWorkingScore + CDbl((CLng(main.SQLCon.Recordset![nummat1]) + lngMatilda1) / 100)
    dblWorkingScore = dblWorkingScore + CDbl((CLng(main.SQLCon.Recordset![nummat2]) + lngMatilda2) / 50)
    dblWorkingScore = dblWorkingScore + CDbl((CLng(main.SQLCon.Recordset![numsherman]) + lngSherman) / 10)
    dblWorkingScore = dblWorkingScore + CDbl((CLng(main.SQLCon.Recordset![numspit1]) + lngSpitfire1) / 10)
    dblWorkingScore = dblWorkingScore + CDbl((CLng(main.SQLCon.Recordset![numspit2]) + lngSpitfire2) / 5)
    dblWorkingScore = dblWorkingScore + CDbl((CLng(main.SQLCon.Recordset![numspit5]) + lngSpitfire5) / 3)
    dblWorkingScore = dblWorkingScore + CDbl((CLng(main.SQLCon.Recordset![numhurr]) + lngHurricane) / 10)
    dblWorkingScore = dblWorkingScore + CDbl((CLng(main.SQLCon.Recordset![numlancaster]) + lngLancaster) / 1)
    dblWorkingScore = dblWorkingScore + CDbl((CLng(main.SQLCon.Recordset![numhalifax]) + lngHalifax) / 25)
    dblWorkingScore = dblWorkingScore + CDbl((CLng(main.SQLCon.Recordset![numwellington]) + lngWellington) / 17.5)
    dblWorkingScore = dblWorkingScore + CDbl((CLng(main.SQLCon.Recordset![numlandingcraft]) + lngLandingCraft) / 90)
    dblWorkingScore = dblWorkingScore + CDbl((CLng(main.SQLCon.Recordset![numlandingship]) + lngLandingShip) / 8)
    dblWorkingScore = dblWorkingScore + CDbl((CLng(main.SQLCon.Recordset![numfrigates]) + lngFrigates) / 50)
    dblWorkingScore = dblWorkingScore + CDbl((CLng(main.SQLCon.Recordset![numbattleship]) + lngBattleShip) / 5)
    dblWorkingScore = dblWorkingScore + CDbl((CLng(main.SQLCon.Recordset![numcarrier]) + lngcarrier) / 1.5)
    
    dblWorkingScore = dblWorkingScore + CDbl(CLng(main.SQLCon.Recordset![numaaguns]) / 100)
    dblWorkingScore = dblWorkingScore + CDbl(CLng(main.SQLCon.Recordset![numpilboxes]) / 100)
    dblWorkingScore = dblWorkingScore + CDbl(CLng(main.SQLCon.Recordset![numturrets]) / 100)
    dblWorkingScore = dblWorkingScore + CDbl(CLng(main.SQLCon.Recordset![numballoons]) / 85)
    dblWorkingScore = dblWorkingScore + CDbl(CLng(main.SQLCon.Recordset![numwatermines]) / 100)
    dblWorkingScore = dblWorkingScore + CDbl(CLng(main.SQLCon.Recordset![amountland]) / 15)
    dblWorkingScore = dblWorkingScore + CDbl(CLng(main.SQLCon.Recordset![nummp]) / 30)
    dblWorkingScore = dblWorkingScore + CDbl(CLng(main.SQLCon.Recordset![numspies]) / 3)
    dblWorkingScore = dblWorkingScore + CDbl(CLng(main.SQLCon.Recordset![numexplorers]) / 100)
    dblWorkingScore = dblWorkingScore + CDbl(CLng(main.SQLCon.Recordset![numfarms]) / 10)
    dblWorkingScore = dblWorkingScore + CDbl(CLng(main.SQLCon.Recordset![nummines]) / 10)
    dblWorkingScore = dblWorkingScore + CDbl(CLng(main.SQLCon.Recordset![numrefineries]) / 30)
    dblWorkingScore = dblWorkingScore + CDbl(CLng(main.SQLCon.Recordset![nummills]) / 1)
    
    lngFinalScore = CLng(dblWorkingScore)
    
    main.SQLCon.Recordset![score] = lngFinalScore
    
    main.SQLCon.Recordset.Update
    
    main.SQLCon.Recordset.MoveNext
Loop

main.SQLCon.Recordset.Close
End Sub
