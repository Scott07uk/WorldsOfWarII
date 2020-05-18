<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else

'code goes below here

'define varibles
Dim rsBanner
Dim lngbannerid
Dim lngadshows

if trim(Request.Form("banner")) = "" then
		Response.Redirect("admin_add_banner_form.asp?error=1")
	else
		set rsbanner = server.CreateObject("ADODB.Recordset")
		
		rsBanner.LockType = 3
		
		rsBanner.CursorType = 2
		
		strsql = "SELECT * FROM tblBanners ORDER BY ADID DESC;"
		
		rsBanner.Open strsql, adocon
		
		if rsBanner.EOF then
				lngbannerid = 1
				lngadshows = 0
			else
				lngbannerid = clng(rsbanner("Adid")) + 1
				lngadshows = clng(rsbanner("adshows"))
		end if
	
		with rsbanner
			.AddNew
			.Fields("adid") = lngbannerid
			.Fields("adtext") = trim(Request.Form("banner"))
			.Fields("adshows") = lngadshows
			.Fields("addedby") = strusername
			.fields("addedon") = now()
			.Fields("active") = 1
			
			.Update
		end with
		
		rsBanner.Close 
		
		set rsbanner = nothing
		
		Response.Redirect("admin.asp?msg=19")
end if
		
		

'code goes above here

end if
%>
<!---#include file="includes\end_check.asp"--->