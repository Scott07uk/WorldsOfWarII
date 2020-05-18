Attribute VB_Name = "Backup"
Public Sub performBackup(intTickNo As Integer)

'varivble declerations
Dim strAbuseReportsFile As String
Dim strAccountsFile As String
Dim strActiveEmailsFile As String
Dim strAnnouncementsFile As String
Dim strAttackResultsFile As String
Dim strBannedEamilFile As String
Dim strBannersFile As String
Dim strCheetingReportsFile As String
Dim strConfigFile As String
Dim lngNoRecords As Long

Dim rsAbuseReports As TblAbuseReports
Dim rsAccounts As TblAccounts
Dim rsActiveEmails As TblActiveEmails
Dim rsAnnouncements As TblAnnouncements
'ReDim bob As TblAttackResults
'Dim rsBannedEmail As TblBannedEmail
'Dim rsBanners As TblBanners
'Dim rsCheetingReports As TblCheatingReports
'Dim rsConfig As TblConfig

'define the file locations
strAbuseReportsFile = App.Path & "\Backups\" & intTickNo & "tblAbuseReports.bak"
strAccountsFile = App.Path & "\Backups\" & intTickNo & "tblAccounts.bak"
strActiveEmailsFile = App.Path & "\Backups\" & intTickNo & "tblActiveEmails.bak"
strAnnouncementsFile = App.Path & "\Backups\" & intTickNo & "tblAnnouncements.bak"
strAttackResultsFile = App.Path & "\Backups\" & intTickNo & "tblAttackResults.bak"
strBannedEmailFile = App.Path & "\Backups\" & intTickNo & "tblBannedEmail.bak"
strBannersFile = App.Path & "\Backups\" & intTickNo & "tblBanners.bak"
strCheetingReportsFile = App.Path & "\Backups\" & intTickNo & "tblCheetingReports.bak"
strConfigFile = App.Path & "\Backups\" & intTickNo & "tblConfig.bak"


'generate the sql query to get the data out
main.SQLCon.RecordSource = "SELECT * FROM tblabuseReports;"

main.SQLCon.Refresh
lngNoRecords = 1

Open strAbuseReportsFile For Random As #2 Len = Len(rsAbuseReports)

Do While Not main.SQLCon.Recordset.EOF
    rsAbuseReports.Details = Trim(CStr(main.SQLCon.Recordset![Details]))
    rsAbuseReports.Replied = CBool(main.SQLCon.Recordset![Replied])
    rsAbuseReports.ReportedLocation = Trim(CStr(main.SQLCon.Recordset![ReportedLocation]))
    rsAbuseReports.ReportID = CLng(main.SQLCon.Recordset![ReportID])
    rsAbuseReports.Type = Trim(CStr(main.SQLCon.Recordset![Type]))
    rsAbuseReports.UnRead = CBool(main.SQLCon.Recordset![UnRead])
    rsAbuseReports.Username = Trim(CStr(main.SQLCon.Recordset![Username]))

    Put #2, lngNoRecords, rsAbuseReports
    lngNoRecords = lngNoRecords + 1
    main.SQLCon.Recordset.MoveNext
Loop

Close #2

'clean up that table
main.SQLCon.Recordset.Close
lngNoRecords = 1

'generate the sql to backup the accounts table
main.SQLCon.RecordSource = "SELECT * FROM tblaccounts;"

main.SQLCon.Refresh
    
Open strAccountsFile For Random As #2 Len = Len(rsAccounts)

Do While Not main.SQLCon.Recordset.EOF
    rsAccounts.Admin = CBool(main.SQLCon.Recordset![Admin])
    rsAccounts.AmountLand = CLng(main.SQLCon.Recordset![AmountLand])
    rsAccounts.ConAA = CInt(main.SQLCon.Recordset![ConAA])
    rsAccounts.ConBF = CInt(main.SQLCon.Recordset![ConBF])
    rsAccounts.ConDC = CInt(main.SQLCon.Recordset![ConDC])
    rsAccounts.ConHF = CInt(main.SQLCon.Recordset![ConHF])
    rsAccounts.ConHMP = CInt(main.SQLCon.Recordset![ConHMP])
    rsAccounts.ConJF = CInt(main.SQLCon.Recordset![ConJF])
    rsAccounts.ConLF = CInt(main.SQLCon.Recordset![ConLF])
    rsAccounts.ConMF = CInt(main.SQLCon.Recordset![ConMF])
    rsAccounts.ConMP = CInt(main.SQLCon.Recordset![ConMP])
    rsAccounts.ConRF = CInt(main.SQLCon.Recordset![ConRF])
    rsAccounts.ConRS = CInt(main.SQLCon.Recordset![ConRS])
    rsAccounts.ConSC = CInt(main.SQLCon.Recordset![ConSC])
    rsAccounts.ConSF = CInt(main.SQLCon.Recordset![ConSF])
    rsAccounts.ConSPF = CInt(main.SQLCon.Recordset![ConSPF])
    rsAccounts.ConSRC = CInt(main.SQLCon.Recordset![ConSRC])
    rsAccounts.ConTicks = CInt(main.SQLCon.Recordset![ConTicks])
    rsAccounts.Continent = CInt(main.SQLCon.Recordset![Continent])
    rsAccounts.ConWF = CInt(main.SQLCon.Recordset![ConWF])
    rsAccounts.Country = Trim(CStr(main.SQLCon.Recordset![Country]))
    rsAccounts.CountryID = CInt(main.SQLCon.Recordset![CountryID])
    rsAccounts.CountryName = Trim(CStr(main.SQLCon.Recordset![CountryName]))
    rsAccounts.County = Trim(CStr(main.SQLCon.Recordset![County]))
    rsAccounts.CurCon = Trim(CStr(main.SQLCon.Recordset![CurCon]))
    rsAccounts.CurResearch = Trim(CStr(main.SQLCon.Recordset![CurResearch]))
    rsAccounts.DOB = CDate(main.SQLCon.Recordset![DOB])
    rsAccounts.Email = Trim(CStr(main.SQLCon.Recordset![Email]))
    rsAccounts.FirstLogin = CBool(main.SQLCon.Recordset![FirstLogin])
    rsAccounts.FontName = Trim(CStr(main.SQLCon.Recordset![FontName]))
    rsAccounts.FontSize = CInt(main.SQLCon.Recordset![FontSize])
    rsAccounts.Food = CLng(main.SQLCon.Recordset![Food])
    rsAccounts.Gold = CInt(main.SQLCon.Recordset![Gold])
    rsAccounts.Holiday = CBool(main.SQLCon.Recordset![Holiday])
    rsAccounts.IRCNick = Trim(CStr(main.SQLCon.Recordset![IRCNick]))
    rsAccounts.Iron = CLng(main.SQLCon.Recordset![Iron])
    rsAccounts.LastLogin = CInt(main.SQLCon.Recordset![LastLogin])
    rsAccounts.LeaderName = Trim(CStr(main.SQLCon.Recordset![LeaderName]))
    rsAccounts.Money = CLng(main.SQLCon.Recordset![Money])
    rsAccounts.NumAAGun = CLng(main.SQLCon.Recordset![NumAAGuns])
    rsAccounts.NumBalloons = CLng(main.SQLCon.Recordset![NumBalloons])
    rsAccounts.NumBattleship = CInt(main.SQLCon.Recordset![NumBattleship])
    rsAccounts.NumCargoShip = CLng(main.SQLCon.Recordset![NumCargoShip])
    rsAccounts.NumCarrier = CInt(main.SQLCon.Recordset![NumCarrier])
    rsAccounts.NumCommandos = CLng(main.SQLCon.Recordset![NumCommandos])
    rsAccounts.NumExplorers = CLng(main.SQLCon.Recordset![NumExplorers])
    rsAccounts.NumFarms = CLng(main.SQLCon.Recordset![NumFarms])
    rsAccounts.NumFrigates = CInt(main.SQLCon.Recordset![NumFrigates])
    rsAccounts.NumHalifax = CLng(main.SQLCon.Recordset![NumHalifax])
    rsAccounts.NumHolidays = CInt(main.SQLCon.Recordset![NumHolidays])
    rsAccounts.NumHurr = CLng(main.SQLCon.Recordset![NumHurr])
    rsAccounts.NumInfantry = CLng(main.SQLCon.Recordset![NumInfantry])
    rsAccounts.NumJeeps = CLng(main.SQLCon.Recordset![NumJeeps])
    rsAccounts.NumLancaster = CLng(main.SQLCon.Recordset![NumLancaster])
    rsAccounts.NumLandingCraft = CLng(main.SQLCon.Recordset![NumLandingCraft])
    rsAccounts.NumLandingShip = CLng(main.SQLCon.Recordset![NumLandingShip])
    rsAccounts.NumMat1 = CLng(main.SQLCon.Recordset![NumMat1])
    rsAccounts.NumMat2 = CLng(main.SQLCon.Recordset![NumMat2])
    rsAccounts.NumMills = CLng(main.SQLCon.Recordset![NumMills])
    rsAccounts.NumMines = CLng(main.SQLCon.Recordset![NumMines])
    rsAccounts.NumMP = CLng(main.SQLCon.Recordset![NumMP])
    rsAccounts.NumPilboxes = CLng(main.SQLCon.Recordset![NumPilboxes])
    rsAccounts.NumPoW = CLng(main.SQLCon.Recordset![NumPoW])
    rsAccounts.NumRefineries = CLng(main.SQLCon.Recordset![NumRefineries])
    rsAccounts.NumSherman = CLng(main.SQLCon.Recordset![NumSherman])
    rsAccounts.NumSpies = CLng(main.SQLCon.Recordset![NumSpies])
    rsAccounts.NumSpit1 = CLng(main.SQLCon.Recordset![NumSpit1])
    rsAccounts.NumSpit2 = CLng(main.SQLCon.Recordset![NumSpit2])
    rsAccounts.NumSpit5 = CLng(main.SQLCon.Recordset![NumSpit5])
    rsAccounts.NumTurrets = CLng(main.SQLCon.Recordset![NumTurrets])
    rsAccounts.NumVotes = CInt(main.SQLCon.Recordset![NumVotes])
    rsAccounts.NumWaterMines = CLng(main.SQLCon.Recordset![NumWaterMines])
    rsAccounts.NumWellington = CLng(main.SQLCon.Recordset![NumWellington])
    rsAccounts.Password = Trim(CStr(main.SQLCon.Recordset![Password]))
    rsAccounts.PowerBase = CBool(main.SQLCon.Recordset![PowerBase])
    rsAccounts.ProtectionTime = CInt(main.SQLCon.Recordset![ProtectionTime])
    rsAccounts.ResAAS = CInt(main.SQLCon.Recordset![ResAAS])
    rsAccounts.ResAC = CInt(main.SQLCon.Recordset![ResAC])
    rsAccounts.ResAES = CInt(main.SQLCon.Recordset![ResAES])
    rsAccounts.ResBGP = CInt(main.SQLCon.Recordset![ResBGP])
    rsAccounts.ResBLC = CInt(main.SQLCon.Recordset![ResBLC])
    rsAccounts.ResCE = CInt(main.SQLCon.Recordset![ResCE])
    rsAccounts.ResCS = CInt(main.SQLCon.Recordset![ResCS])
    rsAccounts.ResD = CInt(main.SQLCon.Recordset![ResD])
    rsAccounts.ResE = CInt(main.SQLCon.Recordset![ResE])
    rsAccounts.ResearchTicks = CInt(main.SQLCon.Recordset![ResearchTicks])
    rsAccounts.ResEP = CInt(main.SQLCon.Recordset![ResEP])
    rsAccounts.ResFG = CInt(main.SQLCon.Recordset![ResFG])
    rsAccounts.ResFPC = CInt(main.SQLCon.Recordset![ResFPC])
    rsAccounts.ResG = CInt(main.SQLCon.Recordset![ResG])
    rsAccounts.ResLB = CInt(main.SQLCon.Recordset![ResLB])
    rsAccounts.ResMBS = CInt(main.SQLCon.Recordset![ResMBS])
    rsAccounts.ResMPC = CInt(main.SQLCon.Recordset![ResMPC])
    rsAccounts.ResMPV = CInt(main.SQLCon.Recordset![ResMPV])
    rsAccounts.ResMWP = CInt(main.SQLCon.Recordset![ResMWP])
    rsAccounts.ResRP = CInt(main.SQLCon.Recordset![ResRP])
    rsAccounts.ResSBS = CInt(main.SQLCon.Recordset![ResSBS])
    rsAccounts.ResSF = CInt(main.SQLCon.Recordset![ResSF])
    rsAccounts.ResSIT = CInt(main.SQLCon.Recordset![ResSIT])
    rsAccounts.ResUWR = CInt(main.SQLCon.Recordset![ResUWR])
    rsAccounts.ResVGM = CInt(main.SQLCon.Recordset![ResVGM])
    rsAccounts.ResWC = CInt(main.SQLCon.Recordset![ResWC])
    rsAccounts.ResWMP = CInt(main.SQLCon.Recordset![ResWMP])
    rsAccounts.ShowAds = CBool(main.SQLCon.Recordset![ShowAds])
    rsAccounts.Suspended = CBool(main.SQLCon.Recordset![Suspended])
    rsAccounts.TicksMsg = CInt(main.SQLCon.Recordset![TicksMsg])
    rsAccounts.Username = Trim(CStr(main.SQLCon.Recordset![Username]))
    rsAccounts.Vote = Trim(CStr(main.SQLCon.Recordset![Vote]))
    rsAccounts.Wood = CLng(main.SQLCon.Recordset![Wood])
    
    Put #2, lngNoRecords, rsAccounts
    lngNoRecords = lngNoRecords + 1
    main.SQLCon.Recordset.MoveNext
    
Loop

Close #2

'clean up that table
main.SQLCon.Recordset.Close
lngNoRecords = 1

main.SQLCon.RecordSource = "SELECT * FROM tblActiveEmails;"

main.SQLCon.Refresh

Open strActiveEmailsFile For Random As #2 Len = Len(rsActiveEmails)

Do While Not main.SQLCon.Recordset.EOF
    rsActiveEmails.Active = CBool(main.SQLCon.Recordset![Active])
    rsActiveEmails.Code = Trim(main.SQLCon.Recordset![Code])
    rsActiveEmails.Email = Trim(main.SQLCon.Recordset![Email])
    rsActiveEmails.Ticks = CInt(main.SQLCon.Recordset![Ticks])
    
    Put #2, lngNoRecords, rsActiveEmails
    lngNoRecords = lngNoRecords + 1
    main.SQLCon.Recordset.MoveNext
Loop

Close #2
main.SQLCon.Recordset.Close
lngNoRecords = 1

main.SQLCon.RecordSource = "SELECT * FROM tblAnnouncements;"

main.SQLCon.Refresh

Open strAnnouncementsFile For Random As #2 Len = Len(rsAnnouncements)

Do While Not main.SQLCon.Recordset.EOF
    rsAnnouncements.AnnouncementID = CLng(main.SQLCon.Recordset![AnnouncementID])
    rsAnnouncements.Message = main.SQLCon.Recordset![Message]
    rsAnnouncements.Posted = main.SQLCon.Recordset![Posted]
    rsAnnouncements.Username = main.SQLCon.Recordset![Username]
    
    Put #2, lngNoRecords, rsAnnouncements
    lngNoRecords = lngNoRecords + 1
    main.SQLCon.Recordset.MoveNext
Loop

Close #2
main.SQLCon.Recordset.Close
lngNoRecords = 1
    
    
End Sub
