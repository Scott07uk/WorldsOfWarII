<!---#include file="includes/check.asp"--->
<%
'<!---SPELL CHECKED--->
Dim rsVote		'collection use to do the work
Dim strName		'the continent name
Dim strBanner	'the continent banner
Dim strIRC		'the irc channel
Dim strMessage	'the message from commanders
Dim blnPowerBase

'get the user input
strName = Request.Form("name")
strBanner = Request.Form("banner")
strIRC = Request.Form("ircchannel")
strMessage = Request.Form("message")

'validate and format the user input
if trim(strName) = "" then
		strName = "[Worlds of War II]"
end if
strname = server.HTMLEncode(strname)
if trim(strBanner) = "" then
		strBanner = ""
end if
strbanner = server.HTMLEncode(strbanner)
if trim(strIRC) = "" then
		strIRC = ""
end if
strirc = server.HTMLEncode(strirc)
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

strsql = "SELECT powerbase from tblaccounts WHERE username = '" & strusername & "';"

rsVote.Open strsql, adocon

blnPowerBase = cbool(rsvote("powerbase"))

if blnpowerbase = false then
		Response.Redirect("politics.asp")
	else
		rsVote.Close 
		rsVote.LockType = 3
		
		rsVote.CursorType = 2
		
		strsql = "SELECT name, continentno, banner, message, ircchannel FROM tblcontinents WHERE continentno = " & lngcontinentno
		
		rsVote.Open strsql, adocon
		
		if rsVote.EOF then
				rsVote.AddNew
				rsVote.Fields("continentno") = lngcontinentno
		end if
		
		with rsvote
			.Fields("name") = strname
			.Fields("banner") = strbanner
			.Fields("message") = strmessage
			.Fields("ircchannel") = strirc
			
			.Update
		end with
		Response.Redirect("politics.asp?msg=2")
end if

rsVote.Close 

set rsvote = nothing

%>
<!---#include file="includes/end_check.asp"--->