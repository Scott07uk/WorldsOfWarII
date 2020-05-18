<!---#include file="includes/check.asp"--->

<!---SPELL CHECKED--->

'Error code:
'Error 1 - Empty field
'Error 2 - Invalid co-ordinates format
'Error 3 - Co-ordinates do not exist
'Error 4 - Anti-spam activated

<%
If Request.Form("continent") = "" Then
	Response.Redirect("sendmsg.asp?error=1")
Else
	If IsNumeric(Request.Form("continent")) = False Then
		Response.Redirect("sendmsg.asp?error=2")
	Else
		If Request.Form("country") = "" Then
			Response.Redirect("sendmsg.asp?error=1")
		Else
			If IsNumeric(Request.Form("country")) = False Then
				Response.Redirect("sendmsg.asp?error=2")
			Else
				strSQL = "SELECT TicksMsg FROM TblAccounts WHERE Username = '" & strUsername & "';"
				
				Dim rscheck 'Check for existence of target country.
				Set rscheck = Server.CreateObject("ADODB.Recordset")
				rscheck.LockType = 3
				rscheck.CursorType = 2
				rscheck.Open StrSQL, Adocon
				If rscheck("TicksMsg") > 20 Then
					Response.Redirect("msgs.asp?error=4")
					rscheck.Close
					Set rscheck = Nothing
				Else
					if blnTickRunning = False then
							Call rscheck.Update("TicksMsg", CInt(rscheck("TicksMsg") + 1))
					end if
					rscheck.Close
				StrSQL = "SELECT TblAccounts.Continent, CountryID FROM TblAccounts WHERE Continent = " & Request.Form("continent") & " AND CountryID = " & Request.Form("country")
				rscheck.Open StrSQL, Adocon
				If rscheck.EOF Then
					rscheck.Close
					Set rscheck = Nothing
					Response.Redirect("sendmsg.asp?error=3")
				Else
					rscheck.Close
					Set rscheck = Nothing
						If Request.Form("subject") = "" Then
							Response.Redirect("sendmsg.asp?error=1")
						Else
							If Request.Form("body") = "" Then
								Response.Redirect("sendmsg.asp?error=1")
							Else
								Dim rssendmessage 'Send message.
								StrSQL = "SELECT TblMessages.* FROM TblMessages ORDER BY MessageID DESC;"
								Set rssendmessage = Server.CreateObject("ADODB.Recordset")
								rssendmessage.CursorType=2
								rssendmessage.LockType=3
								rssendmessage.Open StrSQL, Adocon

								If rssendmessage.EOF Then
									MsgID = 1
								Else
									MsgID = CLng(rssendmessage("MessageID"))
									MsgID = CLng(MsgID + 1)
								End If

								strSubject = server.htmlencode(request.form("subject"))
								strMsg = trim( request.form("body") )
								strMsg = server.htmlencode(strMsg)
								strMsg = Replace(strMsg,vbCrlf,"<BR>")
								strMsg = Replace(strMsg,"[b]","<b>")
								strMsg = Replace(strMsg,"[B]","<b>")
								strMsg = Replace(strMsg,"[/b]","</b>")
								strMsg = Replace(strMsg,"[/B]","</b>")
								strMsg = Replace(strMsg,"[i]","<i>")
								strMsg = Replace(strMsg,"[I]","<i>")
								strMsg = Replace(strMsg,"[/i]","</i>")
								strMsg = Replace(strMsg,"[/I]","</i>")
								strMsg = Replace(strMsg,"[u]","<u>")
								strMsg = Replace(strMsg,"[U]","<u>")
								strMsg = Replace(strMsg,"[/u]","</u>")
								strMsg = Replace(strMsg,"[/U]","</u>")

								rssendmessage.AddNew
								rssendmessage.Fields("FromContinent") = LngContinentNo
								rssendmessage.Fields("FromCountry") = LngCountryNo
								rssendmessage.Fields("ToContinent") = CLng(Request.Form("continent"))
								rssendmessage.Fields("ToCountry") = CLng(Request.Form("country"))
								rssendmessage.Fields("Subject") = strSubject
								rssendmessage.Fields("Body") = strMsg
								rssendmessage.Fields("MessageID") = MsgID
								rssendmessage.Fields("New") = 1
								rssendmessage.Fields("DateSent") = date()
								rssendmessage.Update

								rssendmessage.Close
								Set rssendmessage = Nothing

								Response.Redirect("msgs.asp?sent=1")
							End If
						End If
					End If
				End If
			End If
		End If
	End If
End If
%>

<!---#include file="includes/end_check.asp"--->