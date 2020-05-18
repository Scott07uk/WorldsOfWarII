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
blnConnections = Request.Form("connectallowed")
blnServices = Request.Form("services")
blnStats = Request.Form("stats")
intStatus = Request.Form("status")

blnconnections = cbool(blnconnections)
blnservices = cbool(blnservices)
blnstats = cbool(blnstats)
intStatus = cint(intstatus)

if (isnumeric(intNumeric) = false) OR (strname = "") OR (stradmin="") OR (stremail="") then
		Response.Redirect("admin_add_irc_form.asp?error=1")
	else
		strsql = "SELECT ServerNumeric, AllowConnect, NetworkBots, StatsBot, OwnedBy, ContactEmail, Address, status FROM tblircservers WHERE servernumeric = " & intNumeric
		
		set rsIRC = server.CreateObject("ADODB.Recordset")
		
		rsIRC.LockType = 3
		rsIRC.CursorType = 2
		
		rsIRC.Open strsql, adocon
		
		if not rsIRC.EOF then
				Response.Redirect("admin_add_irc_form.asp?error=2")
			else
				with rsirc
					.AddNew 
					.Fields("numeric") = intnumeric
					.Fields("address") = strname
					.Fields("ownedby") = stradmin
					.Fields("contactemail") = stremail
					.Fields("allowconnect") = blnconnections
					.Fields("networkbots") = blnservices
					.Fields("statsbot") = blnstats
					.Fields("status") = intstatus
				
					.Update
				end with
				
				Response.Redirect("admin.asp?msg=18")
		end if
		
		rsIRC.Close 
		
		set rsirc = nothing
end if
		
		

'code goes above here

end if
%>
<!---#include file="includes\end_check.asp"--->