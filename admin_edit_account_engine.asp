<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else


'Declair varibles 
Dim rsEdit			'collection to be edited
Dim strEditName		'the username of the account to be edited
Dim strEmail		'the email address
Dim strLeaderEdit	'the leader name of the person
Dim strCountryEdit	'the country name of the account
Dim strIrcNick		'the accounts irc nickname
Dim intTicksMsg		'the number of messages the user has sent this tick
Dim intProTime		'the protectiontime the accou t has left
Dim intHolidays		'the number of holidays the account has
Dim strCounty		'the county the user is from
Dim strCountry		'the country the user is from.  

'get the form sumbission

strEditName = Request.Form("username1")
strEmail = Request.Form("email")
strLeaderEdit = Request.Form("leadername")
strCountryEdit = Request.Form("countryname")
strircnick = Request.Form("ircnick")
intticksmsg = Request.Form("ticksmsg")
intprotime = Request.Form("protectiontime")
intholidays = Request.Form("numholiday")
strcounty = Request.Form("county")
strcountry = Request.Form("country")

'validate the user input
if streditname = "" then
		Response.Redirect("admin.asp")
	else
		if isnumeric(intholidays) = false then
				intholidays = -1
		end if
		
		if isnumeric(intprotime) = false then
				intprotime = -1
		end if
		if isnumeric(intticksmsg) = false then
				intticksmsg = -1
		end if
		
		'generate the sql script
		strSQL = "SELECT email, leadername, countryname, ircnick, country, county, ticksmsg, protectiontime, numholidays FROM tblaccounts WHERE username = '" & streditname & "';"
		
		'create the object to perfrom the update
		
		set rsEdit = server.CreateObject("ADODB.Recordset")
		
		'set the object attributes to allow it to update
		rsEdit.LockType = 3
		rsEdit.CursorType = 2
		
		'open the recordset collection
		rsEdit.open strsql, adocon
		
		if rsEdit.EOF then
				Response.Redirect("admin_edit_account_form1.asp?error=1" & streditname)
			else
				'edit the collection
				if stremail <> "" then
						call rsEdit.Update("email", trim(stremail))
				end if
				
				if trim(strLeaderEdit) <> "" then
						call rsEdit.Update("leadername", trim(strLeaderEdit))
				end if
				
				if trim(strCountryEdit) <> "" then
						call rsEdit.Update("countryname", trim(strCountryEdit))
				end if
				
				call rsEdit.Update("ircnick", trim(strircnick))
				
				if intticksmsg > 0 then
						call rsEdit.Update("ticksmsg", cint(intticksmsg))
				end if
				
				if intprotime > 0 then
						call rsEdit.Update("protectiontime", cint(intprotime))
				end if
				
				if intholidays > -1 then
						call rsEdit.Update("numholidays", cint(intholidays))
				end if
				
				if trim(strcountry) <> "" then
						call rsEdit.Update("country", trim(strcountry))
				end if
				
				if trim(strcounty) <> "" then
						call rsEdit.Update("county", trim(strcountry))
				end if
				
		end if
		
		rsEdit.Close
		
		set rsedit = nothing
		
		Response.Redirect("admin.asp?msg=3")
end if
end if
%>
<!---#include file="includes\end_check.asp"--->