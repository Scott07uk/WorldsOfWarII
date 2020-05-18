<%
'<!---SPELL CHECKED--->
Dim strEmailCode
Dim strEmail
Dim strRedir
Dim adocon
Dim strSQL
Dim rsEmail
Dim strEmailDomain
Dim objEmail

strEmail = Request.Form("email")
strRedir = Request.Form("redir")

if strEmail = "" then
		strRedir = "activate_email1.asp?error=1"
	else
		if InStr(strEmail, "@") = 0 then
				strRedir = "activate_email1.asp?error=2"
			else
				if InStr(strEmail, ".") = 0 then
						strRedir = "activate_email1.asp?error=2"
					else
						Set adoCon = Server.CreateObject("ADODB.Connection")
		
						adocon.ConnectionString = "Provider=SQLOLEDB;Server=neptune;User ID=WoWUsr;Password=password;Database=WoWII;"
		
						adocon.Open 
						
						strsql = "SELECT tblBannedEmail.type FROM tblBannedEmail WHERE Emailaddress = '" & strEmail & "';"
						
						set rsEmail = server.CreateObject("ADODB.Recordset")
						
						rsEmail.LockType = 3
						
						rsEmail.CursorType = 2
						
						rsEmail.Open strsql, adocon
						
						if not rsEmail.EOF then
								strRedir = "activate_email1.asp?error=3"
							else
								rsEmail.Close 
								
								strEmailDomain = right(trim(stremail), clng(len(stremail) - InStr(trim(strEmail), "@")))
								
								strsql = "SELECT tblBannedEmail.type FROM tblBannedEmail WHERE EmailDomain = '" & strEmailDomain & "';"
								
								rsEmail.Open strsql, adocon
								
								if not rsEmail.EOF then
										strRedir = "activate_email.asp?error=4"
									else
										rsEmail.Close
										
										randomize
										strEmailCode = clng(rnd(1000)*1000)
										randomize
										strEmailCode = StrEmailCode & clng(rnd(10000)*10000)
										randomize
										stremailCode = stremailcode & clng(rnd(100000)*100000)
										
										stremailcode = replace(stremailcode, ".", "5")
										stremailcode = trim(stremailcode)
										
										strsql = "SELECT tblActiveEmails.* FROM tblactiveEmails WHERE email = '" & stremail & "';"
										
										rsEmail.Open strsql, adocon
										
										if not rsEmail.EOF then
												strRedir = "activate_email1.asp?error=5"
											else
												if strRedir = "" then
														strRedir = "add_email2.asp"
												end if
												with rsemail
													.AddNew
													.Fields("email") = stremail
													.Fields("active") = 0
													.Fields("code") = stremailcode
													.Fields("ticks") = 0
													.Update
												end with
												
												set objEmail = Server.CreateObject("JMail.SMTPMail")
												
												with objEmail
													.serverAddress = "neptune.home.local"
													.sender = "No-Reply@worldsofwar.co.uk"
													.AddRecipient stremail
													.Subject = "Worlds of War II Email Activation"
													.body = "Thank you for activating your email address with Worlds of War II.  To add your email to the active addresses database you will need the information below.  " & vbcrlf & vbcrlf & "If you did not want your email address to be activated then please disregard this message, your email address will not be stored.  " & vbcrlf & vbcrlf & "The code you need to contine your registration is: " & vbcrlf & vbcrlf & stremailcode & vbcrlf & vbcrlf & "If you have closed the signup window, go to http://www.worldsofwar.co.uk and click on the register link, then activate email.  "
													
													.execute
												end with
												
												set objemail = nothing
										end if
								end if
						end if
						
						rsEmail.Close
						adocon.Close
						
						set rsemail = nothing
						set adocon = nothing
				end if
		end if
end if
		
Response.Redirect(strRedir)
		
%>
		
		
								
								