<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else

'codes goes below here

if isdate(Request.Form("deldate")) = false then
		Response.Redirect("admin_del_olf_msg_form1.asp?error=1")
	else
		Dim rsMessage
		
		strSQL = "SELECT * FROM  tblMessages;"
		
		set rsMessage = server.CreateObject("ADODB.Recordset")
		
		rsMessage.LockType = 3
		
		rsMessage.CursorType = 2
		
		rsMessage.Open strsql, adocon
		
		do while not rsMessage.EOF
			if rsMessage("datesent") < cdate(Request.Form("deldate")) then
					rsMessage.Delete 
					Response.Write("1")
			end if
			rsMessage.MoveNext 
		loop
		
		rsMessage.Close 
		
		set rsmessage = nothing
		
		Response.Redirect("admin.asp?msg=8")
end if
'code goes above here


end if
%>
<!---#include file="includes\end_check.asp"--->