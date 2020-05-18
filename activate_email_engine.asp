<%
'<!---SPELL CHECKED--->
Dim adocon
Dim strSQL
Dim rsemail
Dim strEmail
Dim strcode
Dim objEmail

stremail = Request.Form("email")
strcode = Request.Form("code")

if (stremail = "") OR (strcode = "") then
		Response.Redirect("add_email2.asp?error=1")
	else
		set adocon = server.CreateObject("ADODB.Connection")
		
		adocon.ConnectionString = "Provider=SQLOLEDB;Server=neptune;User ID=WoWUsr;Password=password;Database=WoWII;"
		
		adocon.Open
		
		set rsemail = server.CreateObject("ADODB.Recordset")
		
		rsemail.LockType = 3
		
		rsemail.CursorType = 2
		
		if strcode = "" then
				strsql = "SELECT code, ticks FROM tblactiveemails WHERE email = '" & stremail & "';"
				
				rsemail.Open strsql, adocon
				
				if rsemail.EOF then
						Response.Redirect("add_email.asp?error=2")
					else
						set objemail = server.CreateObject("JMail.SMTPMail")
						
						with objemail
						
							.serverAddress = "neptune.home.local:25"
							.sender = "Support@worldsofwar.co.uk"
							.AddRecipient stremail
							.Subject = "Worlds of War II Email Activation"
							.body = "Thank you for activating your email address with Worlds of War II.  To add your email to the active addresses database you will need the information below.  " & vbcrlf & vbcrlf & "If you did not want your email address to be activated then please disregard this message, your email address will not be stored.  " & vbcrlf & vbcrlf & "The code you need to contine your registration is: " & vbcrlf & vbcrlf & rsemail("code") & vbcrlf & vbcrlf & "If you have closed the signup window, go to http://www.worldsofwar.co.uk and click on the register link, then activate email.  "
													
							.execute
						end with
						
						call rsemail.Update("ticks", cint(0))
				end if
			else
				strsql = "SELECT code, ticks, active FROM tblactiveemails WHERE email = '" & stremail & "';"
				
				rsemail.Open strsql, adocon
				
				if rsemail.EOF then
						Response.Redirect("add_email.asp?error=2")
					else
						if rsemail("code") = strcode then
								call rsemail.Update("ticks", cint(0))
								call rsemail.Update("active", cint(1))
								Response.Redirect("add_email_done.asp")
							else
								Response.Redirect("add_email.asp?error=3")
						end if
				end if
		end if
		
		rsemail.Close
		
		set rsemail = nothing
		
		adocon.Close
		
		set adocon = nothing
end if

%>
		