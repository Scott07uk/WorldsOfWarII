<!---#include file="tick_control.asp"--->
<%
'Declair the varibles to be used in this page
Dim strUsername		'holds the username the user entered
Dim strPassword		'holds the password the user entered
Dim adoCon			'holes the connection to the sql Server
Dim rsLogins		'holds the recorset to validate the logins
Dim rsRecord		'holds the collection which is used to save the login
Dim strSQL			'holds the sql used to query the sql server

'get the data the user entered in the login form
strUsername = Request.Form("username")
strPassword = Request.Form("password")


'make sure the user has entered a username to login with
if strUsername = "" then
		'redirecrt the user back to the login page
		Response.Redirect("login.asp")
	else
		'the user has entered a username it now needs to be checked.  
		
		'set the object to connect to the sql server
		Set adoCon = Server.CreateObject("ADODB.Connection")

		'create the connection string for the sql server
		strCon = "Provider=SQLOLEDB;Server=neptune;User ID=WoWUsr;Password=password;Database=WoWII;"

		'apply the connection string to the connection object
		adoCon.connectionstring = strCon

		'open the connection to the sql server
		adocon.Open 

		'create the recordset object to hold the login data
		Set rsLogins = Server.CreateObject("ADODB.Recordset")
		set rsRecord = Server.CreateObject("ADODB.Recordset")
		
		'create the sql to get the users password out of the database
		strSQL = "SELECT tblaccounts.password, tblaccounts.firstlogin, tblaccounts.suspended, tblaccounts.lastlogin FROM tblaccounts WHERE username = '" & strUsername & "';"
		
		'set the recordset proerpties so they can be updated
		rsLogins.LockType = 3
		rsRecord.LockType = 3
		
		rsLogins.CursorType = 2
		rsRecord.CursorType = 2
		
		'open the recordset from the sql server
		
		rsLogins.Open strSQL, adoCon
		
		'make sure there are some records to check
		if rsLogins.EOF then
				'there is no macthing username
				Response.Redirect("login.asp?error=1")
			else
				'validate the users password
				if strPassword <> rsLogins("password") then
						'the user has entered the wrong password
						Response.Redirect("login.asp?error=1")
					else
						'make sure the user has an active account
						if cbool(rsLogins("suspended")) = true then
								'the user account has been suspended
								Response.Redirect("login.asp?error=2")
							else
										
								'the user is clear the go in
								
								'save a cookie to the users computer to keep them loged in
								Response.Cookies("WoWII")("username") = strUsername
								Response.Cookies("WoWII")("password") = strPassword
								Response.Cookies("WOWII").Expires = now() + 0.5	'allow the cookie to stay on the users machine fot 7 days
								'update the users account so we know the last time they loged in
								if blnTickRunning = false then
										call rsLogins.Update ("lastLogin","0")
								end if
								'record the login
								strsql = "SELECT tblLogins.* FROM tbllogins ORDER by loginid DESC;"
								
								rsRecord.Open strsql, adocon
								
								Dim lngNextItem
								
								if rsRecord.EOF then
										lngNextItem = 1
									else
										lngNextItem = clng(rsrecord("loginid"))+1
								end if
								
								rsRecord.AddNew 
								rsRecord.Fields("loginid") = lngNextItem
								rsRecord.Fields("username") = strusername
								rsRecord.Fields("ip") = Request.ServerVariables("REMOTE_ADDR")
								rsRecord.Fields("logintime") = now()
								
								rsRecord.Update 
								
								rsRecord.Close 
								'check to see if the user has loged in before
								if cbool(rsLogins("firstlogin")) = true then
										'mark the user has having loged in for the first time
										'rsLogins.Fields("firstlogin") = 0
										call rsLogins.Update("firstlogin", "0")
										'redirect them to the intro movie
										Response.Redirect("film_story.asp")
									else
										'redirect the user to the main game
										Response.Redirect("main.asp")
								end if

						end if
				end if
		end if
		
		'close the connection to the sql server
		
		rsLogins.Close
		
		adoCon.Close
		
		'kill the objects
		set rsLogins = nothing
		set rsaccount = nothing
		set adocon = nothing
end if
%>
								