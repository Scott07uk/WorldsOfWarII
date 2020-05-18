<!---#include file="includes\check.asp"--->

<!---SPELL CHECKED--->

<%

Dim rsspies
Dim lngReportID

'get the report id
lngReportid = Request.QueryString("id")

'validate the input
if isnumeric(lngReportid) = false then
		Response.Redirect("spies.asp")
	else
		'create the objects to handle the operation
		set rsspies = server.CreateObject("ADODB.Recordset")
		
		'set the object attributes
		rsspies.LockType = 3 
		
		rsspies.CursorType = 2
		
		strsql = "SELECT * FROM TblIntelegenceReports WHERE username = '" & strusername & "' AND reportid = " & lngreportid
		
		rsspies.Open strsql, adocon
		
		if rsspies.EOF then
				Response.Redirect("spies.asp")
			else
				rsspies.Delete 
				
				Response.Redirect("spies.asp?error=12")
		end if
		
		rsspies.Close 
		
		set rsspies = nothing
end if

%>
<!---#include file="includes\end_check.asp"--->