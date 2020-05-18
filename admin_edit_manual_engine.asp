<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else

'code goes below here

'define varibles
Dim rsManual
Dim intid
Dim strTitle
Dim strManual
Dim intorder
Dim blnsidebar

'get the input from the user
intid = Request.Form("id")
strtitle = Request.Form("title")
strManual = Request.Form("entry")
intorder = Request.Form("order")
blnsidebar = Request.Form("onsidebar")

blnsidebar = cbool(blnsidebar)


if (isnumeric(intid) = false) OR (isnumeric(intorder) = false) then
		Response.Redirect("admin_change_manual_form.asp")
	else
		strsql = "SELECT entry, entryname, onsidebar, ordering FROM tblmanual WHERE enteryid = " & intid
		
		set rsmanual = server.CreateObject("ADODB.Recordset")
		
		rsmanual.LockType = 3
		rsmanual.CursorType = 2
		
		rsmanual.Open strsql, adocon
		
		if rsmanual.EOF then
				Response.Redirect("admin_change_manual_form.asp")
			else
				with rsmanual
					.Fields("entry") = trim(strmanual)
					.Fields("entryname") = trim(strtitle)
					.Fields("onsidebar") = blnsidebar
					.Fields("ordering") = intorder
					
					.Update
				end with
				
				Response.Redirect("admin.asp?msg=21")
		end if
		
		rsmanual.Close 
		
		set rsmanual = nothing
end if
		
		

'code goes above here

end if
%>
<!---#include file="includes\end_check.asp"--->