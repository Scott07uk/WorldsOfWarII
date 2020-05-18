<%
'SPELL CHECKED

option Explicit

'decalir varibles
Dim strsql
Dim adocon
Dim lngcontinentno
Dim lngcountryno
Dim rssignup
Dim stremail
Dim strusername
Dim strpass1
Dim strpass2
Dim strcountryname
Dim strleadername
Dim strfromcounty
Dim strfromcountry
Dim dtedob
Dim strircnick
Dim strRedir

'get the users input from the form
strusername = Request.Form("username")
strpass1 = Request.Form("pass1")
strpass2 = Request.Form("pass2")
stremail = Request.Form("email")
strfromcounty = Request.Form("county")
strfromcountry = Request.Form("country")
dtedob = Request.Form("dobday") & "/" & Request.Form("dobmonth") & "/" & Request.Form("dobyear")
strircnick = Request.Form("ircnick")
strcountryname = Request.Form("countryname")
strleadername = Request.Form("leadername")

'create a redirection url so the user does not have to enter every thing in again
strRedir = "username=" & strusername & "&stremail=" & stremail & "&fromcountry=" & strfromcountry & "&fromcounty=" & strfromcounty & "&ircnick=" & strircnick & "&countryname=" & strcountryname & "&leadername=" & strleadername


'validate that the user has entered required fields
if strusername = "" then
		Response.Redirect("register2.asp?error=1&" & strRedir)
	else
		'make sure they entered an email
		if stremail = "" then
				Response.Redirect("register2.asp?error=2&" & strredir)
			else
				'make sure they entered a country name
				if strcountryname = "" then
						Response.Redirect("register2.asp?error=3&" & strredir)
					else
						if strleadername = "" then
								Response.Redirect("register2.asp?error=4&" & strredir)
							else
								'check the two passwords are the same
								if strpass1 <> strpass2 then
										Response.Redirect("register2.asp?error=5&" & strredir)
									else
										'make sure the date of birth is valid
										if isnumeric(Request.Form("dobyear")) = false then
												Response.Redirect("register2.asp?error=6&" & strredir)
											else
												if clng(trim(Request.Form("dobyear"))) > year(now()) then
														Response.Redirect("register2.asp?error=6&" & strredir)
													else
														Dim intmonthdays
														select case cint(trim(Request.Form("dobmonth")))
															case 1
																intmonthdays = 31
															case 2
																intmonthdays = 29
															case 3
																intmonthdays = 31
															case 4
																intmonthdays = 30
															case 5
																intmonthdays = 31
															case 6
																intmonthdays = 30
															case 7
																intmonthdays = 31
															case 8
																intmonthdays = 31
															case 9
																intmonthdays = 30
															case 10
																intmonthdays = 31
															case 11
																intmonthdays = 30
															case 12
																intmonthdays = 31
														end select
														
														if cint(Request.Form("dobday")) > intmonthdays then
																Response.Redirect("register2.asp?error=6&" & strredir)
															else
																'check to see if the email address is active
																set adocon = server.CreateObject("ADODB.Connection")
																
																adocon.ConnectionString = "Provider=SQLOLEDB;Server=neptune;User ID=WoWUsr;Password=password;Database=WoWII;"
																
																adocon.open
																
																set rssignup = server.CreateObject("ADODB.recordset")
																
																rssignup.LockType = 3
																
																rssignup.CursorType = 2
																
																strsql = "SELECT tblactiveemails.active FROM tblactiveemails WHERE email = '" & stremail & "';"
																
																rssignup.Open strsql, adocon
																
																if rssignup.EOF then
																		'the email address is not in the active addresses database
																		Response.Redirect("register2.asp?error=7&" & strredir)
																	else
																		if cbool(rssignup("active")) = false then
																				'the email address is in the database but it has not been activated
																				Response.Redirect("register2.asp?error=8&" & strredir)
																			else
																				'the email is good
																				'check to see if the usernaem has already been taken
																				rssignup.Close
																				
																				strsql = "SELECT tblaccounts.username FROM tblaccounts where username = '" & strusername & "';"
																				
																				rssignup.Open strsql, adocon
																				
																				if rssignup.EOF = false then
																						'the username is already in use
																						Response.Redirect("register2.asp?error=9&" & strredir)
																					else
																						'the username is good, calculate the users new location.  
																						rssignup.Close
																						
																						strsql = "SELECT tblcontinents.continentno, numaccounts FROM tblcontinents WHERE newAccounts = 1 AND fullup = 0 ORDER by numaccounts ASC;"
																						
																						rssignup.Open strsql, adocon
																						
																						'check to see if there is no continents avaible
																						if rssignup.EOF then
																								rssignup.Close
																								
																								strsql = "SELECT continentno, name, message, fullup, newaccounts, numaccounts FROM tblcontinents ORDER by continentno DESC;"
																								
																								rssignup.Open strsql, adocon
																								
																								if rssignup.EOF then
																										'the world is totally empty
																										lngcontinentno = 1
																									else
																										lngcontinentno = clng(clng(rssignup("continentno")) + 1)
																								end if
																								
																								with rssignup
																									.AddNew
																									.Fields("continentno") = lngcontinentno
																									.Fields("name") = "[Worlds of War II]"
																									.Fields("message") = "No Message has been set yet"
																									.Fields("fullup") = 0
																									.Fields("newaccounts") = 1 
																									.Fields("numaccounts") = cint(1)
																									.Update
																								end with
																								
																								lngcountryno = 1
																							else
																								lngcontinentno = rssignup("continentno")
																								call rssignup.Update("numaccounts", cint(rssignup("numaccounts")+1))
																								rssignup.Close
																								lngcountryno = 0
																								
																								strsql = "SELECT tblaccounts.countryid FROM tblaccounts WHERE continent = " & lngcontinentno & " ORDER BY countryid ASC;"
																								
																								rssignup.Open strsql, adocon
																								
																								'check to see if the all the accounts in this continent have been deleted
																								if rssignup.EOF then
																										lngcountryno = 1
																									else
																										Dim intcounter
																										intcounter = 1
																										'we need to loop though all the accounts until we see one thats free
																										Dim blnexitvar
																										blnexitvar = cbool(false)
																										do while blnexitvar = false
																											if rssignup.EOF then
																													lngcountryno = intcounter
																													blnexitvar = true
																												else
																													if rssignup("countryid") = intcounter then
																															'do nothing as there is a valid account in this space
																														else
																															'there is an account missing in the middle of a continent
																															lngcountryno = intcounter
																															blnexitvar = true
																													end if
																													intcounter = intcounter + 1
																											end if
																											
																											if intcounter = 16 then
																													blnexitvar = true
																													lngcountryno = 0
																											end if
																											if not rssignup.EOF then
																													rssignup.MoveNext	
																											end if
																										loop
																										'if the countryno is still 0 then an error has occoured and the continet is full
																										'create a new continent for the account
																										
																										if lngcountryno = 0 then
																												rssignup.Close
																										
																												strsql = "SELECT continentno, name, message, fullup, newaccounts, numaccounts FROM tblcontinents ORDER by continentno DESC;"
																										
																												rssignup.Open strsql, adocon
																										
																												'make sure things really havent screwed up
																												if rssignup.EOF then
																														lngcontinentno = 1
																													else
																														lngcontinentno = rssignup("continentno") + 1
																												end if
																										
																												'add a record for the continent
																												with rssignup
																													.AddNew
																													.Fields("continentno") = lngcontinentno
																													.Fields("name") = "[Worlds of War II]"
																													.Fields("message") = "No Message has been set yet"
																													.Fields("fullup") = 0
																													.Fields("newaccounts") = 1 
																													.Fields("numaccounts") = cint(1)
																													.Update
																												end with
																										
																												lngcountryno = 1
																											else
																												if lngCountryno = 15 then
																														rssignup.Close 
																														
																														strsql = "SELECT fullup FROM tblcontinents WHERE continentno = " & lngcontinentno
																														
																														rssignup.Open strsql, adocon
																														call rssignup.Update("fullup", cbool(true))
																												end if
																										end if
																								end if
																						end if
																						
																						'time to create the account
																						rssignup.Close
																						
																						strsql = "SELECT username, password, email, county, country, countryname, leadername, dob, continent, countryid, ircnick from tblaccounts;"
																						
																						with rssignup
																							.Open strsql, adocon
																							
																							.AddNew
																							.Fields("username") = strusername
																							.Fields("password") = strpass1
																							.Fields("email") = stremail
																							.Fields("county") = trim(strfromcounty)
																							.Fields("country") = trim(strfromcountry)
																							.Fields("countryname") = trim(strcountryname)
																							.Fields("leadername") = trim(strleadername)
																							.Fields("dob") = dtedob
																							.Fields("continent") = lngcontinentno
																							.Fields("countryid") = lngcountryno
																							.Fields("ircnick") = strircnick
																							
																							.Update
																							
																						end with
																						
																						rssignup.Close 
																						
																						strsql = "SELECT tblconfig.Signupmsg FROM tblconfig;"
																						
																						rssignup.Open strsql, adocon
																						
																						if rssignup.EOF then
																								'do nothing
																							else
																								Dim objEmail
																								set objEmail = Server.CreateObject("JMail.SMTPMail")
																								with objEmail
																									.serverAddress = "127.0.0.1"
																									.sender = "Support@worldsofwar.co.uk"
																									.AddRecipient stremail
																									.Subject = "Worlds of War II Registration"
																									.body = rssignup("signupmsg")
																								
																									.execute
																								end with
																						end if
																						
																						Response.Redirect("register_done.asp")
																				end if
																		end if
																end if
																
																rssignup.Close
																adocon.Close
																set rssignup = nothing
																set adocon = nothing
														end if
												end if
										end if
								end if
						end if
				end if
		end if
end if

%>