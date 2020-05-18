<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else

'code goes below here

Dim rsVote		'collection use to do the work
Dim strName		'the continent name
Dim strBanner	'the continent banner
Dim strIRC		'the irc channel
Dim strMessage	'the message from commanders
Dim blnlocked

'get the user input
strName = Request.Form("name")
strBanner = Request.Form("banner")
strIRC = Request.Form("ircchannel")
strMessage = Request.Form("message")
blnLocked = cbool(Request.Form("locked"))

'validate and format the user input
if trim(strName) = "" then
		strName = "[Worlds of War II]"
end if
if trim(strBanner) = "" then
		strBanner = ""
end if
if trim(strIRC) = "" then
		strIRC = ""
end if
if trim(strMessage) = "" then
		strMessage="No message has been set.  "
end if

strMessage = trim(server.HTMLEncode(strmessage))

strmessage = replace(strmessage,vbcrlf,"<BR>")
strmessage = replace(strmessage,"[b]","<B>")
strmessage = replace(strmessage,"[B]","<B>")
strmessage = replace(strmessage,"[/b]","</B>")
strmessage = replace(strmessage,"[/B]","<B>")
strmessage = replace(strmessage,"[i]","<I>")
strmessage = replace(strmessage,"[I]","<I>")
strmessage = replace(strmessage,"[/i]","</I>")
strmessage = replace(strmessage,"[/I]","</I>")
strmessage = replace(strmessage,"[u]","<U>")
strmessage = replace(strmessage,"[U]","<U>")
strmessage = replace(strmessage,"[/u]","</U>")
strmessage = replace(strmessage,"[/U]","</U>")

'create the objects to do the work
set rsVote = server.CreateObject("ADODB.Recordset")

rsVote.LockType = 3
		
rsVote.CursorType = 2
		
strsql = "SELECT name, continentno, banner, message, ircchannel FROM tblcontinents WHERE continentno = " & Request.Form("continent")
		
rsVote.Open strsql, adocon
		
if rsVote.EOF then
		Response.Redirect("admin_con_settings_form1.asp")
	else
		
		with rsvote
			.Fields("name") = strname
			.Fields("banner") = strbanner
			.Fields("message") = strmessage
			.Fields("ircchannel") = strirc
			
			.Update
		end with
		Response.Redirect("admin.asp?msg=11")
end if

rsVote.Close 

set rsvote = nothing

'code goes above here
end if
%>
<!---#include file="includes\end_check.asp"--->