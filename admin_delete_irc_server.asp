<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else

'code goes below here

'define varibles to be used
Dim rsIRC
Dim intNumeric

intNumeric = Request.QueryString("numeric")

if isnumeric(intnumeric) = false then
		Response.Redirect("admin_irc_list.asp")
	else
		set rsIRC = server.CreateObject("ADODB.Recordset")
		
		rsIRC.LockType = 3
		rsIRC.CursorType = 2
		
		strsql = "SELECT * FROM tblircservers WHERE numeric = " & intnumeric
		
		rsIRC.Open strsql, adocon
		
		if rsIRC.EOF then
				Response.Redirect("admin_irc_list.asp")
			else
				rsIRC.Delete 
				
				Response.Redirect("admin.asp?msg=17")
		end if
		
		rsIRC.Close 
		
		set rsirc = nothing
end if

'code goes above here

end if
%>
<!---#include file="includes\end_check.asp"--->