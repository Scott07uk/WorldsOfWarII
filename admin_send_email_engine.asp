<%
server.ScriptTimeout = 120
%>
<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
on error resume next
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
'code goes below here

Dim rsAddress
Dim objEmail
Dim strMessage
Dim lngEmail
Dim strSubject

strMessage = Request.Form("message")
strSubject = Request.Form("subject")

if strmessage = "" then
		Response.Redirect("admin_email_all_users.asp?error=1")
	else
		if strsubject = "" then
				Response.Redirect("admin_email_all_users.asp?error=1")
			else
				'create the objects
				
				set rsAddress = server.CreateObject("ADODB.Recordset")
				
				set objEmail = server.CreateObject("Jmail.SMTPMail")
				
				strsql = "SELECT email from tblaccounts;"
				
				rsAddress.Open strsql, adocon
				
				do while not rsAddress.EOF 
					with objemail 
						.ServerAddress = "127.0.0.1"
						.Sender = "Support@worldsofwar.co.uk"
						.AddRecipient rsaddress("email")
						.Subject = strsubject
						.HTMLBody = strmessage
						
						.Execute 
						
						.ClearRecipients 
						
						rsAddress.MoveNext 
					end with
				loop
				
				rsAddress.Close 
				
				set objemail = nothing
				set rsaddress = nothing
				
				Response.Redirect("admin.asp?error=24")
		end if
end if

'code goes above here
end if
%>
<!---#include file="includes\end_check.asp"--->