<%
option explicit
%>
<!---#include file="includes\check.asp"--->

<%
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else

'code goes below here

Dim strmsg
Dim strdbfield
Dim rsMsg
select case Request.form("type")
	case "creator"
		strdbfield = "creatorsmessage"
	case "register"
		strdbfield = "signupmsg"
end select

strmsg = Request.Form("message")

strsql = "SELECT " & strdbfield & " FROM tblconfig;"

set rsmsg = server.CreateObject("ADODB.Recordset")

rsMsg.LockType = 3 
rsMsg.Cursortype = 2 

rsMsg.Open strsql, adocon

call rsMsg.Update(strdbfield, strmsg)

rsMsg.Update 

rsMsg.Close 

set rsmsg = nothing

Response.Redirect("admin.asp?msg=12")

'code goes above here
end if
%>
<!---#include file="includes\end_check.asp"--->