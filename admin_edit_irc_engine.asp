<%
option explicit
%>
<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else

'code goes below here

'define varibles
Dim rsIRC
Dim intNumeric
Dim strName
Dim strAdmin
Dim strEmail
Dim blnConnections
Dim blnServices
Dim blnStats
Dim intStatus

'get the input from the user
intNumeric = Request.Form("numeric")
strname = Request.Form("name")
stradmin = Request.Form("adminname")
strEmail = Request.Form("adminEmail")
blnConnections = Request.Form("connectallow")
blnServices = Request.Form("services")
blnStats = Request.Form("stats")
intStatus = Request.Form("status")

blnconnections = cbool(blnconnections)
blnservices = cbool(blnservices)
blnstats = cbool(blnstats)
intStatus = cint(intstatus)

if isnumeric(intNumeric) = false then
		Response.Redirect("admin_irc_list.asp")
	else
		strsql = "SELECT ServerNumeric, AllowConnect, NetworkBots, StatsBot, OwnedBy, ContactEmail, Address, status FROM tblircservers WHERE servernumeric = " & intNumeric
		
		set rsIRC = server.CreateObject("ADODB.Recordset")
		
		rsIRC.LockType = 3
		rsIRC.CursorType = 2
		
		rsIRC.Open strsql, adocon
		
		if rsIRC.EOF then
				Response.Redirect("admin_irc_list.asp")
			else
				if strname <> "" then
						rsIRC.Fields("address") = strname
				end if
				if stradmin <> "" then
						rsIRC.Fields("ownedby") = stradmin
				end if
				if strEmail <> "" then
						rsIRC.Fields("contactemail") = stremail
				end if
				rsIRC.Fields("allowconnect") = blnconnections
				rsIRC.Fields("networkbots") = blnservices
				rsIRC.Fields("statsbot") = blnstats
				rsIRC.Fields("status") = intstatus
				
				rsIRC.Update
				
				Response.Redirect("admin.asp?msg=16")
		end if
		
		rsIRC.Close 
		
		set rsirc = nothing
end if
		
		

'code goes above here

end if
%>
<!---#include file="includes\end_check.asp"--->