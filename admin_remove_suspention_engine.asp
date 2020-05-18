<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else

'code goes below here

Dim rsSuspend			'recordset collection to remove the suspension
Dim objEmail			'object to email the user

if Request.QueryString("username") = "" then
		Response.Redirect("admin_remove_suspention_form.asp")
	else
		'create the sql script
		strSQL = "SELECT suspended, email FROM tblaccounts WHERE username = '" & Request.QueryString("username") & "';"
		
		'create the object and set the properties
		set rsSuspend = server.CreateObject("ADODB.Recordset")
		
		rsSuspend.LockType = 3
		rsSuspend.CursorType = 2
		
		'open the recordset
		rsSuspend.Open strsql, adocon
		
		'check to make sure the username is valid
		if rsSuspend.EOF then
				Response.Redirect("admin_remove_suspention_form.asp")
			else
				call rsSuspend.Update("suspended", cbool(False))
				
				set objEmail = server.CreateObject("JMail.SMTPMail")
				
				with objEmail
					.ServerAddress = "127.0.0.1"
					.Sender = "Support@worldsofwar.co.uk"
					.AddRecipient rsSuspend("email")
					.Subject = "Worlds of War II Suspension Removed"
					.Body = "Your account on Worlds of War II with username " & Request.QueryString("username") & " is now no longer suspended. You may log in and play the game again.  " & vbcrlf & vbcrlf & "Worlds of War Creators."
					.Execute 
				end with
				
				set objemail = nothing
				Response.Redirect("admin.asp?msg=2")
		end if
		
		rsSuspend.Close
		
		set rsSuspend = nothing
end if
				

'code goes above here

end if
%>
<!---#include file="includes\end_check.asp"--->