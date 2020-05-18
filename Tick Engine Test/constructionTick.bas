Attribute VB_Name = "constructionTick"
Public Sub ConTicks()

'open da database
Dim strCurCon As String
main.SQLCon.RecordSource = "SELECT tblaccounts.username, Conticks, curCon, conAA, conJF, conMF, conSF, conSPF, conHMP, conLF, conHF, conWF, conBF, conRF, conMP, conDC, conRS, conSC, conSRC FROM tblaccounts"
main.SQLCon.Refresh

Set rsConstruction = main.SQLCon.Recordset
    Print #1, "+-----------------------------------------+" & vbCrLf
    Print #1, "|   CONSTRUCTION PROCESS AND NEWS SENT!   |" & vbCrLf
    Print #1, "+-----------------------------------------+" & vbCrLf & vbCrLf
    Do While Not rsConstruction.EOF
        If rsConstruction![ConTicks] <> 0 Then
            rsConstruction![ConTicks] = rsConstruction![ConTicks] - 1
            If rsConstruction![ConTicks] = 0 Then
                strCurCon = Trim(rsConstruction![CurCon])
                rsConstruction![CurCon] = ""
                Select Case strCurCon
                    Case "AA"
                        rsConstruction![ConAA] = 3
                    Case "JF"
                        rsConstruction![ConJF] = 3
                    Case "MF"
                        rsConstruction![ConMF] = 3
                    Case "SF"
                        rsConstruction![ConSF] = 3
                    Case "SPF"
                        rsConstruction![ConSPF] = 3
                    Case "HMP"
                        rsConstruction![ConHMP] = 3
                    Case "LF"
                        rsConstruction![ConLF] = 3
                    Case "HF"
                        rsConstruction![ConHF] = 3
                    Case "WF"
                        rsConstruction![ConWF] = 3
                    Case "BF"
                        rsConstruction![ConBF] = 3
                    Case "RF"
                        rsConstruction![ConRF] = 3
                    Case "MP"
                        rsConstruction![ConMP] = 3
                    Case "DC"
                        rsConstruction![ConDC] = 3
                    Case "RS"
                        rsConstruction![ConRS] = 3
                    Case "SC"
                        rsConstruction![ConSC] = 3
                    Case "SRC"
                        rsConstruction![ConSRC] = 3
                End Select
                main.SQLCon2.RecordSource = "SELECT tblnews.* FROM tblnews ORDER BY newsID DESC;"
                main.SQLCon2.Refresh

                Set rsnews = main.SQLCon2.Recordset
                    intnews = rsnews![NewsID]
                
                rsnews.AddNew
                rsnews![NewsID] = intnews + 1
                rsnews![username] = rsConstruction![username]
                rsnews![TimeMade] = Now()
                rsnews![Message] = "The Construction has been completed"
                rsnews![UnRead] = 1
                rsnews.Update
                rsnews.Close
            End If
            rsConstruction.Update

            Print #1, rsConstruction![username] & " has had Construction processed"
        End If
    rsConstruction.MoveNext
    Loop
rsConstruction.Close

End Sub

