<!---#include file="includes/check.asp"--->

<!---SPELL CHECKED--->

<%
Dim rssettings 'Open account settings details.
Set rssettings = Server.CreateObject("ADODB.Recordset")
StrSQL = "SELECT TblAccounts.Password, Email, County, Country, Dob, Holiday, NumHolidays, IRCNick FROM TblAccounts WHERE Continent = " & LngContinentNo & " AND CountryID = " & LngCountryNo
rssettings.CursorType = 2
rssettings.LockType = 3
rssettings.Open strSQL, Adocon

If Request.Form("oldpassword") <> "" And Request.Form("newpassword") <> "" And Request.Form("confirmnew") <> "" Then
	If rssettings("Password") <> Request.Form("oldpassword") Then
		Response.Redirect("settings.asp?error=1")
	Else
		If Request.Form("newpassword") <> Request.Form("confirmnew") Then
			Response.Redirect("settings.asp?error=2")
		Else
			rssettings.Fields("Password") = Request.Form("newpassword")
		End If
	End If
End If
If rssettings.Fields("Email") <> Request.Form("email") Then
	Dim rsemail 'Check for valid email address.
	Set rsemail = Server.CreateObject("ADODB.Recordset")
	StrSQL = "SELECT TblActiveEmails.Email, Active FROM TblActiveEmails WHERE Email = '" & Request.Form("email") & "' AND Active = 1;"
	rsemail.Open strSQL, Adocon
	If rsemail.EOF Then
		Response.Redirect("settings.asp?error=4")
	Else
		rssettings.Fields("Email") = Request.Form("email")
	End If
	rsemail.Close
	Set rsemail = Nothing
End If
rssettings.Fields("County") = Request.Form("county")
rssettings.Fields("Country") = Request.Form("country")
If Request.Form("dateofbirth") <> "" Then
	If IsDate(Request.Form("dateofbirth")) = False Then
		Response.Redirect("settings.asp?error=5")
	Else
		rssettings.Fields("Dob") = Request.Form("dateofbirth")
	End If
End If
If Request.Form("holiday") = "yes" Then
	If cbool(rssettings.Fields("Holiday")) = False Then
		If rssettings.Fields("NumHolidays") < 2 Then
			rssettings.Fields("Holiday") = cbool(True)
			rssettings.Fields("NumHolidays") = rssettings.Fields("NumHolidays") + 1
		Else
			Response.Redirect("settings.asp?error=3")
		End If
	Else
		rssettings.Fields("Holiday") = cbool(False)
	End If
End If
rssettings.Fields("IRCNick") = Request.Form("ircnick")
rssettings.Update

rssettings.Close
Set rssettings = Nothing

Response.Redirect("settings.asp?msg=1")
%>

<!---#include file="includes/end_check.asp"--->