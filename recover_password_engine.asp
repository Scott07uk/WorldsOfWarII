<%
'SPELL CHECKED

'declair varibles
Dim strMessage			'varible to hold the email message
Dim adocon				'varible to hold the connection to the sql server
Dim strSQL				'varible to hold the sql string
Dim rsRecover			'holds the recordset collection object
Dim strEmail			'holds the email address the user has submitted
Dim objEmail			'holds the email object

'set the message to be emailed to the user, [username] and [password] will be subsititued
strMessage = "Hi," & vbcrlf & "You have requested your Worlds of War II login details. To log in to Worlds of War II, you need to go to http://www.worldsofwar.co.uk and use the login box at the top of the screen. Your login details are below:" & vbcrlf & vbcrlf & "Username: [username]" & vbcrlf & "Password: [password]" & vbcrlf & vbcrlf & "If you are having other problems with Worlds of War II please let us know."

'get the email address
strEmail = Request.Form("email")

'tidy up the user input
strEmail = trim(strEmail)

'validate the user input
if stremail = "" then
		Response.Redirect("login.asp?error=4")
	else
		'create the objects to open the databse
		set adocon = server.CreateObject("ADODB.Connection")
		
		set rsRecover = server.CreateObject("ADODB.Recordset")
		
		'set the sql to query the database
		strsql = "SELECT username, password FROM tblaccounts WHERE email = '" & stremail & "';"
		
		'open the connection to the database server
		adocon.ConnectionString = "Provider=SQLOLEDB;Server=neptune;User ID=WoWUsr;Password=password;Database=WoWII;"
		
		adocon.Open
		
		'open the recordset
		rsRecover.Open strsql, adocon
		
		'check to make sure a record exits
		if rsRecover.EOF then
				'redirect the user to the main page
				Response.Redirect("login.asp?error=5")
			else
				'finalise the message
				strmessage = replace(strmessage, "[username]", rsrecover("username"))
				strmessage = replace(strmessage, "[password]", rsrecover("password"))
				
				set objEmail = Server.CreateObject("JMail.SMTPMail")
				
				with objEmail
					.serveraddress = "127.0.0.1"
					.Sender = "Support@worldsofwar.co.uk"
					.AddRecipient stremail
					.Subject = "Worlds of War II Login Details"
					.Body = strmessage
					.Execute 
				end with
				
				Response.Redirect("login.asp?error=6")
		end if
		
		rsRecover.Close
		adocon.Close 
		
		set rsrecover = nothing
		set adocon = nothing
end if
				

%>