<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
'code goes below here

Dim lngid
Dim rsManual

'get the userinput
lngid = Request.QueryString("id")

if isnumeric(lngid) = false then
		Response.Redirect("admin_change_manual_form.asp")
	else
		strsql = "SELECT * FROM tblmanual WHERE enteryid = " & lngid
		
		set rsmanual = server.CreateObject("ADODB.Recordset")
		
		rsManual.LockType = 3
		
		rsManual.Cursortype =2 
		
		rsManual.Open strsql, adocon
		
		if rsManual.EOF then
				Response.Redirect("admin_change_manual_form.asp")
			else
				rsManual.Delete 
				Response.Redirect("Admin.asp?msg=22")
		end if
		
		rsManual.Close 
		
		set rsmanual = nothing
end if

'code goes above here
end if
%>
<!---#include file="includes\end_check.asp"--->