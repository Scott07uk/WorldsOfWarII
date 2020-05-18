<%
option explicit
%>
<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else

'code goes below here

Dim rsEmail
Dim lngRecordNo

if (trim(Request.Form("domain"))="") AND (trim(Request.Form("email"))="") then
		Response.Redirect("admin_ban_email.asp?error=1")
	else
		if (trim(Request.Form("domain"))<>"") AND (trim(Request.Form("email"))<>"") then
				Response.Redirect("admin_ban_email.asp?error=2")
			else
				set rsemail = server.CreateObject("ADODB.Recordset")
				
				rsEmail.LockType = 3
				
				rsEmail.Cursortype = 2
				
				strsql = "SELECT * FROM tblbannedemail ORDER BY emailid DESC;"
				
				rsEmail.Open strsql, adocon
				
				if rsEmail.EOF then
						lngrecordno = 1
					else
						lngrecordno = clng(rsemail("emailid")) + 1
				end if
				
				with rsemail
					.AddNew
					.Fields("emailid") = lngrecordno
					.Fields("emaildomain") = trim(Request.Form("domain"))
					.Fields("emailaddress") = trim(Request.Form("email"))
					.Fields("type") = 1 
					
					.Update
				end with
				
				Response.Redirect("admin.asp?msg=14")
				
				rsEmail.Close 
				
				set rsemail = nothing
		end if
end if
		

'code goes above here
end if
%>
<!---#include file="includes\end_check.asp"--->