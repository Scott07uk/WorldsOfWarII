<!---#include file="../tick_control.asp"--->
<%
'check file to make sure that the user is loged in
Dim strUserName			'varible to hold the username
Dim strPassword			'varible to hold the password
Dim adocon				'holds the connection to the sql server
Dim rsLogin				'holds the record collection for the logins
Dim strSQL				'holds the sql connection string
Dim lngCountryNo		'the country number of the account
Dim lngContinentNo		'the continent number of the account
Dim strCountryName		'the country name of the account
Dim strLeaderName		'the leader name of the account
Dim lngMoney			'the amount of money the user has
Dim lngFood				'the amount of food the user has
Dim lngWood				'the amount of wood the user has
Dim lngIron				'the amount of iron the user has
Dim lngGold				'the amount of gold the user has
dim lngLand				'the amount of LAN' the guy has
Dim blnAdmin			'holds if the account is admin or not
Dim blnAdsOn			'holds the value of the banner ad
Dim lngScore			'holds the users game score


'get the cookie data
strusername = Request.Cookies("WoWII")("username")
strpassword = Request.Cookies("WoWII")("password")

'create an object to connect to the sql server
Set adoCon = Server.CreateObject("ADODB.Connection")

'prvide the server details
adocon.ConnectionString = "Provider=SQLOLEDB;Server=neptune;User ID=WoWUsr;Password=password;Database=WoWII;"

'privide the sql to query the server with
strSQL = "SELECT password, continent, countryid, leadername, amountland, countryname, money, food, wood, gold, iron, suspended, admin, showads, score FROM tblaccounts WHERE username = '" & strUsername & "';"

'create the recordset object
set rsLogin = server.CreateObject("ADODB.Recordset")

'open the connection to the server
adocon.Open

'open the recordset
rsLogin.Open strsql, adocon

'check to see if the username is valid
if rsLogin.EOF then
		'direct the user to the login page
		if Request.ServerVariables("HTTP_USER-AGENT") = "Mediapartners-Google/2.1" THEN
				Server.Transfer("google" & Request.ServerVariables("SCRIPT_NAME"))
			ELSE
				Response.Redirect("login.asp")
		END IF
		
		'close the recordset
		rsLogin.Close 
	else
		'the username is valid check the password
		if rslogin("password") <> strpassword then
				'the passwords are wrong redirect the user to the login page unless its the google bot
				Response.Redirect("login.asp")
				
				'close the recordset
				rsLogin.Close 
			else
				'the password is correct check to see if the user has been suspended
				if cbool(rslogin("suspended")) = true then
						'the user has been susppended
						Response.Redirect("login.asp?error=2")
						
						'close the recordset
						rsLogin.close
					else
						'the user is clear to play the game so load the varibles
						lngCountryNo = rslogin("CountryID")
						lngContinentNo = rslogin("continent")
						strCountryName = rslogin("countryname")
						strLeaderName = rslogin("leadername")
						lngMoney = clng(rslogin("money"))
						lngFood = clng(rslogin("food"))
						lngWood = clng(rslogin("wood"))
						lngIron = clng(rslogin("iron"))
						lngLAND = clng(rslogin("amountland"))
						lngGold = clng(rslogin("gold"))
						blnAdmin = cbool(rslogin("admin"))
						blnAdsOn = cbool(rslogin("showads"))
						lngScore = clng(rslogin("score"))
						'close the recordset
						rsLogin.Close
						
						
						%>