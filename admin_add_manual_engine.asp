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
strtitle = Request.Form("title")
strManual = Request.Form("entry")
intorder = Request.Form("order")
blnsidebar = Request.Form("onsidebar")

blnsidebar = cbool(blnsidebar)


if (isnumeric(intid) = false) OR (isnumeric(intorder) = false) then
		Response.Redirect("admin_change_manual_form.asp")
	else
		strsql = "SELECT * FROM tblmanual ORDER BY enteryid DESC;"
		
		set rsmanual = server.CreateObject("ADODB.Recordset")
		
		rsmanual.LockType = 3
		rsmanual.CursorType = 2
		
		rsmanual.Open strsql, adocon
		
		if rsmanual.EOF then
				intid = 1
			else
				intid = clng(rsmanual("enteryid")) + 1
		end if 
		
		with rsmanual
			.AddNew 
			.Fields("enteryid") = intid
			.Fields("entry") = trim(strmanual)
			.Fields("entryname") = trim(strtitle)
			.Fields("onsidebar") = blnsidebar
			.Fields("ordering") = intorder
			.Fields("totalrating") = 0
			.Fields("numrating") = 0					
			
			.Update
		end with
		
		rsmanual.Close 
		
		set rsmanual = nothing
		Response.Redirect("admin.asp?msg=23")
end if
		
		

'code goes above here

end if
%>
<!---#include file="includes\end_check.asp"--->