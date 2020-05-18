<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else

'code goes below here

'define varibles to be used
Dim rsbanner
Dim intid

intid = Request.QueryString("id")

if isnumeric(intid) = false then
		Response.Redirect("admin_list_banners.asp")
	else
		set rsbanner = server.CreateObject("ADODB.Recordset")
		
		rsbanner.LockType = 3
		rsbanner.CursorType = 2
		
		strsql = "SELECT * FROM tblbanners WHERE adid=" & clng(intid)
		
		rsbanner.Open strsql, adocon
		
		if rsbanner.EOF then
				Response.Redirect("admin_list_banners.asp")
			else
				rsbanner.Delete 
				
				Response.Redirect("admin.asp?msg=20")
		end if
		
		rsbanner.Close 
		
		set rsbanner = nothing
end if

'code goes above here

end if
%>
<!---#include file="includes\end_check.asp"--->