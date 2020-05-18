Attribute VB_Name = "researchTick"
Public Sub ResearchTicks()

'open da database
Dim strCurRes As String
main.SQLCon.RecordSource = "SELECT tblaccounts.username, researchticks, curResearch, resSF, resCE, resVGM, resBGP, resMPV, resMWP, resMPC, resFPC, resWC, resRP, resUWR, resG, resEP, resLB, resSIT, resBLC, resSBS, resMBS, resAAS, resAES, resCS, resFG, resWMP, resE, resD, resAC FROM tblaccounts"
main.SQLCon.Refresh

Set rsresearch = main.SQLCon.Recordset
    Print #1, "+-----------------------------------------+" & vbCrLf
    Print #1, "|     RESEARCH PROCESS AND NEWS SENT!     |" & vbCrLf
    Print #1, "+-----------------------------------------+" & vbCrLf & vbCrLf
    Do While Not rsresearch.EOF
        If rsresearch![ResearchTicks] > 0 Then
            rsresearch![ResearchTicks] = rsresearch![ResearchTicks] - 1
            If rsresearch![ResearchTicks] = 0 Then
                strCurRes = Trim(rsresearch![CurResearch])
                rsresearch![CurResearch] = ""
                Select Case strCurRes
                    Case "SF"
                        rsresearch![ResSF] = 3
                    Case "CE"
                        rsresearch![ResCE] = 3
                    Case "VGM"
                        rsresearch![ResVGM] = 3
                    Case "BGP"
                        rsresearch![ResBGP] = 3
                    Case "MPV"
                        rsresearch![ResMPV] = 3
                    Case "MWP"
                        rsresearch![ResMWP] = 3
                    Case "MPC"
                        rsresearch![ResMPC] = 3
                    Case "FPC"
                        rsresearch![ResFPC] = 3
                    Case "WC"
                        rsresearch![ResWC] = 3
                    Case "RP"
                        rsresearch![ResRP] = 3
                    Case "UWR"
                        rsresearch![ResUWR] = 3
                    Case "G"
                        rsresearch![ResG] = 3
                    Case "EP"
                        rsresearch![ResEP] = 3
                    Case "LB"
                        rsresearch![ResLB] = 3
                    Case "SIT"
                        rsresearch![ResSIT] = 3
                    Case "BLC"
                        rsresearch![ResBLC] = 3
                    Case "SBS"
                        rsresearch![ResSBS] = 3
                    Case "MBS"
                        rsresearch![ResMBS] = 3
                    Case "AAS"
                        rsresearch![ResAAS] = 3
                    Case "AES"
                        rsresearch![ResAES] = 3
                    Case "CS"
                        rsresearch![ResCS] = 3
                    Case "FG"
                        rsresearch![ResFG] = 3
                    Case "WMP"
                        rsresearch![ResWMP] = 3
                    Case "E"
                        rsresearch![ResE] = 3
                    Case "D"
                        rsresearch![ResD] = 3
                    Case "AC"
                        rsresearch![ResAC] = 3
                End Select
                main.SQLCon2.RecordSource = "SELECT tblnews.* FROM tblnews ORDER BY NewsID DESC;"
                main.SQLCon2.Refresh

                Set rsnews = main.SQLCon2.Recordset
                If rsnews.EOF Then
                        intnews = 1
                    Else
                        intnews = rsnews![NewsID]
                End If

                rsnews.AddNew
                rsnews![NewsID] = intnews + 1
                rsnews![username] = rsresearch![username]
                rsnews![TimeMade] = Now()
                rsnews![Message] = "The research has been completed"
                rsnews![UnRead] = 1
                rsnews.Update
                rsnews.Close
            End If
            rsresearch.Update

            Print #1, rsresearch![username] & " has had research processed"
        End If
    rsresearch.MoveNext
    Loop
rsresearch.Close

End Sub
