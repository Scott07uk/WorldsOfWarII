<!---#include file="includes\check.asp"--->

<!---SPELL CHECKED--->

<%
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
'code goes below here

Dim rspost
Dim strMessage
Dim strSubject
Dim lngOrderNo

'get the user input

strmessage = Request.Form("message")
strsubject = Request.Form("subject")

if (strmessage = "") OR (strsubject = "") then
		Response.Redirect("postannouncement_form.asp?error=1")
	else
		set rspost = server.CreateObject("ADODB.recordset")
		
		rspost.LockType = 3
		
		rspost.CursorType = 2 
		
		strsql = "SELECT * FROM tblannouncements order by announcementid DESC;"
		
		rspost.Open strsql, adocon
		
		if rspost.EOF then
				lngOrderno = 1
			else
				lngorderno = clng(rspost("announcementid")) + 1
		end if
		
		with rspost 
			.AddNew 
			.Fields("announcementid") = lngorderno
			.Fields("message") = strmessage
			.Fields("username") = strusername
			.Fields("posted") = now()
			
			.Update
		end with
		
		
		Response.Redirect("admin.asp?error=")
		
		
		rspost.Close 
		
		set rspost = nothing
end if

'code goes above here
end if
%>
<!---#include file="includes\end_check.asp"--->