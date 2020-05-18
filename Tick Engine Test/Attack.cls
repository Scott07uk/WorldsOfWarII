VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Attack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'attacking units
Private attInfantrySent As Long
Private attCommandoesSent As Long
Private attJeepsSent As Long
Private attMatilda1Sent As Long
Private attMatilda2Sent As Long
Private attShermanSent As Long
Private attSpitfire1Sent As Long
Private attSpitfire2Sent As Long
Private attSpitfire5Sent As Long
Private attHurrincaneSent As Long
Private attHalifaxSent As Long
Private attWellingtonSent As Long
Private attLancasterSent As Long
Private attLandingCraftSent As Long
Private attLandingShipSent As Long
Private attFrigateSent As Long
Private attBattleShipSent As Long
Private attCarrierSent As Long
Private attInfantryLost As Long
Private attCommandoesLost As Long
Private attJeepsLost As Long
Private attMatilda1Lost As Long
Private attMatilda2Lost As Long
Private attShermanLost As Long
Private attSpitfire1Lost As Long
Private attSpitfire2Lost As Long
Private attSpitfire5Lost As Long
Private attHurrincaneLost As Long
Private attHalifaxLost As Long
Private attWellingtonLost As Long
Private attLancasterLost As Long
Private attLandingCraftLost As Long
Private attLandingShipLost As Long
Private attFrigateLost As Long
Private attBattleShipLost As Long
Private attCarrierLost As Long
'defending units
Private defInfantrySent As Long
Private defCommandoesSent As Long
Private defJeepsSent As Long
Private defMatilda1Sent As Long
Private defMatilda2Sent As Long
Private defShermanSent As Long
Private defSpitfire1Sent As Long
Private defSpitfire2Sent As Long
Private defSpitfire5Sent As Long
Private defHurrincaneSent As Long
Private defHalifaxSent As Long
Private defWellingtonSent As Long
Private defLancasterSent As Long
Private defLandingCraftSent As Long
Private defLandingShipSent As Long
Private defFrigateSent As Long
Private defBattleShipSent As Long
Private defCarrierSent As Long
Private defInfantryLost As Long
Private defCommandoesLost As Long
Private defJeepsLost As Long
Private defMatilda1Lost As Long
Private defMatilda2Lost As Long
Private defShermanLost As Long
Private defSpitfire1Lost As Long
Private defSpitfire2Lost As Long
Private defSpitfire5Lost As Long
Private defHurrincaneLost As Long
Private defHalifaxLost As Long
Private defWellingtonLost As Long
Private defLancasterLost As Long
Private defLandingCraftLost As Long
Private defLandingShipLost As Long
Private defFrigateLost As Long
Private defBattleShipLost As Long
Private defCarrierLost As Long
'the static units
Private defPilboxSent As Long
Private defAAGunSent As Long
Private defBalloonSent As Long
Private defWaterMineSent As Long
Private defTurretSent As Long
Private defPilboxLost As Long
Private defAAGunLost As Long
Private defBalloonLost As Long
Private defWaterMineLost As Long
Private defTurretLost As Long
'the industrial units
Private farmsSent As Long
Private minesSent As Long
Private millsSent As Long
Private refineriesSent As Long
Private farmsLost As Long
Private minesLost As Long
Private millsLost As Long
Private refineriesLost As Long
'the control varibles
Private attSendArr(18, 25) As Long
Private attUsernameArr(25) As String
Private defSendArr(23, 25) As Long
Private defUsernameArr(25) As String
Private attIDArr(25) As Long
Private defIDArr(25) As Long
Private landGainArr(25) As Integer
Private attReportIDArr(25) As Long
Private defReportIDArr(25) As Long
Private defCounter As Integer
Private attCounter As Integer
Private blnHomeDefence As Boolean
'set up the objects needed to connect to the database
Private adoCon As Object
Private strSQL As String

Private Sub Class_Initialize()
'this sets up the object
Set adoCon = CreateObject("ADODB.Connection")
adoCon.ConnectionString = strConnectionString
adoCon.CommandTimeout = 15
adoCon.Open
blnHomeDefence = False
defCounter = 0
attCounter = 0
End Sub

Public Sub addDefence(strUsername As String, movementID As Long, lngInfantry As Long, lngCommandoes As Long, lngJeeps As Long, lngMatilda1 As Long, lngMatilda2 As Long, lngSherman As Long, lngSpitfire1 As Long, lngSpitfire2 As Long, lngSpitfire5 As Long, lngHurricnae As Long, lngHalifax As Long, lngWellington As Long, lngLancaster As Long, lngLandingCraft As Long, lngLandingShip As Long, lngFrigates As Long, lngBattleships As Long, lngCarriers As Long)
'This adds some defence to the
If blnHomeDefence = True Then
        'new defence can not be added as the home defence has been added
    Else
        If defCounter > 22 Then
                'there is no room for this attack
            Else
                'add the defence
                defIDArr(defCounter) = movementID
                defSendArr(0, defCounter) = lngInfantry
                definfantry = definfantry + lngInfantry
                defSendArr(1, defCounter) = lngCommandoes
                defCommandoesSent = defCommandoesSent + lngCommandoes
                defSendArr(2, defCounter) = lngJeeps
                defJeepsSent = defJeepsSent + lngJeeps
                defSendArr(3, defCounter) = lngMatilda1
                defMatilda1Sent = defMatilda1Sent = lngMatilda1
                defSendArr(4, defCounter) = lngMatilda2
                defMatilda2Sent = defMatilda2Sent + lngMatilda2
                defsendarr(5,defcounter)
                
        End If
End If
End Sub

Public Sub addAttack()
'This adds an attacker to the attack
End Sub

Public Sub addHome()
'This adds the home defence to the attack
blnHomeDefence = True
End Sub
