<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else

'code goes below here

Dim rsMessage

'generate the sql script
strSQl = "SELECT tblmessages.* FROM tblmessages WHERE messageid = " & Request.QueryString("msg")

set rsMessage = server.CreateObject("ADODB.Recordset")

'set the attributes
rsMessage.LockType = 3

rsMessage.CursorType = 2

rsMessage.Open strsql, adocon

if not rsMessage.EOF then
		rsMessage.Delete 
end if

rsMessage.Close

set rsmessage = nothing

Response.Redirect("admin.asp?msg=7")

'code goes above here

end if
%>
<!---#include file="includes\end_check.asp"--->