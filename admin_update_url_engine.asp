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

Dim rsConfig

strsql = "SELECT gameaddress, portaladdress, forumsaddress, manualaddress, ircsite FROM tblconfig;"

set rsconfig = server.CreateObject("ADODB.Recordset")

rsconfig.LockType = 3 
rsconfig.Cursortype = 2 

rsconfig.Open strsql, adocon

rsConfig.Fields("gameaddress") = trim(Request.Form("game"))
rsConfig.Fields("portaladdress") = trim(Request.Form("portal"))
rsConfig.Fields("forumsaddress") = trim(Request.Form("forums"))
rsConfig.Fields("manualaddress") = trim(Request.Form("manual"))
rsConfig.Fields("ircsite") = trim(Request.Form("irc"))

rsconfig.Update 

rsconfig.Close 

set rsconfig = nothing

Response.Redirect("admin.asp?msg=13")

'code goes above here
end if
%>
<!---#include file="includes\end_check.asp"--->