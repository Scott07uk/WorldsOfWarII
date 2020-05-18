<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else

'code goes below here

'todo list
'Check the account is valid									DONE
'check the new location is valid							DONE
'create a new continent if needed							DONE
'update the new continent details							DONE
'move the account											DONE
'set the votes to 1 and the account voting for themselfs	DONE
'update the old continent details							DONE
'move the incomming/outgoing movements						DONE
'move the messages											DONE
'save some news to tell the user they have been moved		DONE

'Declair the varibles to be used
Dim lngNewcontinent		'the new location of the account
Dim lngNewCountry		'the new location of the account
Dim lngOldContinent		'the old location of the account
Dim lngOldCountry		'the old location of the account
Dim strMoveName			'the username of the account to be moved
Dim rsMove				'the recordset used for the move
Dim intError			'the varibled used to hold the error codes
Dim blnUseLoc			'varible to dedice on user input
Dim blnNewCon			'varible to decide if there is a new continent or not
Dim strNewsMessage		'varible to hold the new news message

'get the input from the user
lngNewContinent = Request.Form("tocontinent")
lngNewcountry = Request.Form("tocountry")
lngOldContinent = Request.Form("fromcontinent")
lngoldCountry = Request.Form("fromcountry")
strmovename = Request.Form("username")
interror = 0


strNewsMessage = "Your account has been relocated to [newcontinent]:[newcountry] by [admin]."

'validate the user input
if strMoveName = "" then
		if (lngoldCountry = "") or (lngoldContinent = "") then
				intError=1
			else
				'we are using the location
				blnUseLoc = True
				'validate to make sure the numbers are correct
				if (isnumeric(lngoldcountry)=False) or (isnumeric(lngoldcontinent)=False) then
						interror = 2
					else
						'the submission is fine
				end if
		end if
	else
		if (lngoldCountry <> "") or (lngoldContinent <> "") then
				intError = 3
			else
				blnUseLoc=False
		end if
end if

'check to see if there were any errors
if intError <> 0 then
		'there was
		Response.Redirect("admin_relocate_form1.asp?error=" & interror)
	else
		'the sumbission was fine
		
		'validate the new location
		if (isnumeric(lngnewcontinent) = false) OR (isnumeric(lngnewcountry) = false) then
				'the locaton is not numeric
				Response.Redirect("admin_relocate_form1.asp?error=4")
			else
				'make sure the country number is within range
				if (lngnewcontinent < 1) OR (lngnewcountry > 15) then
						'there is an error with the range
						Response.Redirect("admin_relocate_form1.asp?error=5")
					else
						'get the username and the location of the account
						strsql = "SELECT username, countryid, continent FROM tblaccounts WHERE "
						
						'generate the where statement based on user input
						if blnUseLoc = True then
								strsql = strsql & "continent = " & lngoldcontinent & " AND countryid = " & lngoldcountry
							else
								strsql = strsql & "username = '" & strmovename & "';"
						end if
						
						'create the objects to perform the move
						set rsMove = server.CreateObject("ADODB.Recordset")
						
						'set the attributes to allow updating
						rsMove.LockType = 3
						rsMove.CursorType = 2
						
						'open the collection to begin updating
						rsMove.Open strsql, adocon
						
						if rsMove.EOF then
								'the account does not exists
								Response.Redirect("admin_relocate_form1.asp?error=6")
							else
								'retrieve all the needed account attributes to local memory
								strMoveName = rsMove("username")
								lngoldcontinent = rsmove("continent")
								lngOldCountry = rsmove("countryid")
								'close the object and check to make sure the move to location is free
								rsMove.Close
								'generate sql to check the space for the movement
								strsql = "SELECT username FROM tblaccounts WHERE continent = " & lngnewcontinent & " AND countryid = " & lngnewcountry
								
								rsMove.Open strsql, adocon
								if not rsMove.EOF then
										'there is already an account in that space
										Response.Redirect("admin_relocate_form1.asp?error=7")
									else
										'close the db and open up the continent to update it
										rsMove.Close
										
										'open the attributed to append/create new continent
										strsql = "SELECT continentno, name, message, fullup, newaccounts, numaccounts FROM tblcontinents WHERE continentno = " & lngnewcontinent
										
										rsMove.Open strsql, adocon
										'check to see if the continent exists or not
										if rsMove.EOF then
												'we need to create a new continent for the user
												blnNewCon = True
												
												'create a new continent for the movement
												With rsMove
													.AddNew
													.Fields("continentno") = lngnewcontinent
													.Fields("Name") = "[Worlds of War II]"
													.Fields("Message") = "No message has been set"
													.Fields("fullup") = cbool(false)
													.Fields("newaccount") = cbool(true)
													.Fields("numaccounts") = 1
													'commence the changes
													.Update
												end with
												
												'new continent saved
											else
												blnNewCon = False
												'save the continent data that there is a new accoiunt
												if rsmove("numaccounts") = 14 then
														'the continent is now full
														call rsMove.Update("fullup", cbool(true))
												end if
												'update nukbr of accounts attribute
												call rsMove.Update("numaccounts", cint(rsmove("Numaccounts")) + 1)
										end if
										rsMove.Close 
										
										'generate the sql to move the account to the new location
										strsql = "SELECT vote, username, numvotes, continent, countryid, powerbase FROM tblaccounts WHERE username = '" & strmovename & "';"
										
										rsMove.Open strsql, adocon
										
										'move the account (use a multiple edit update to allow for the composit key in the accounts table)
										rsMove.Fields("continent") = lngnewcontinent
										rsMove.Fields("countryid") = lngnewcountry
										
										'update the power base election details
										rsMove.Fields("vote") = rsmove("username")
										rsMove.Fields("numvotes") = 1
										
										'set the power base attribute to the correct value
										rsMove.Fields("powerbase") = cbool(blnNewCon)
										rsMove.Update
										
										rsMove.Close
										
										'update the old continent so the space can be filled
										strsql = "SELECT numaccounts, fullup FROM tblcontinents WHERE continentno = " & lngoldcontinent
										
										rsMove.Open strsql, adocon
										
										'catch any errors
										if rsMove.EOF then
												'do nothing
											else
												'update the continent
												call rsMove.Update("fullup", cbool(false))
												call rsMove.Update("numaccounts", cint(rsmove("numaccounts")) - 1)
										end if
										
										rsMove.Close 
										
										'generate the sql script to edit the incomming units
										strsql = "SELECT tocontinent, tocountry FROM tblmovements WHERE tocontinent = " & lngoldcontinent & " AND tocountry = " & lngoldcountry
										
										rsMove.Open strsql, adocon
										
										'relocate all the movements
										do while not rsMove.EOF 
											call rsMove.Update("tocontinent", lngnewcontinent)
											call rsMove.Update("tocountry", lngnewcountry)
											
											'move to the next extry
											rsMove.MoveNext 
										loop
										
										'close the collection
										rsMove.Close
										
										'generate the sql to move all the outgoing units
										strsql = "SELECT fromcontinent, fromcountry FROM tblmovements WHERE fromcontinent = " & lngoldcontinent & " AND fromcountry = " & lngoldcountry
										
										rsMove.Open strsql, adocon
										
										'relocate all the movements
										do while not rsMove.EOF
											call rsMove.Update("fromcontinent", lngnewcontinent)
											call rsMove.Update("fromcountry", lngnewcountry)
											
											'move to the next entery
											rsMove.MoveNext
										loop
										
										'close the recordset
										rsMove.Close
										
										'create the sql to move the messages to the account to the new location
										
										strsql = "SELECT tocontinent, tocountry FROM tblmessages WHERE tocontinent = " & lngoldcontinent & " AND tocountry = " & lngoldcountry
										
										rsMove.Open strsql, adocon
										
										'relocate all the messages to the account
										do while not rsMove.EOF 
											call rsMove.Update("tocontinent", lngnewcontinent)
											call rsMove.Update("tocountry", lngnewcountry)
											
											'move to the next extry
											rsMove.MoveNext 
										loop
										
										'close the collection
										rsMove.Close
										
										'generate the sql to move all the outgoing messages
										strsql = "SELECT fromcontinent, fromcountry FROM tblmessages WHERE fromcontinent = " & lngoldcontinent & " AND fromcountry = " & lngoldcountry
										
										rsMove.Open strsql, adocon
										
										'relocate all the messages in te collection
										do while not rsMove.EOF
											call rsMove.Update("fromcontinent", lngnewcontinent)
											call rsMove.Update("fromcountry", lngnewcountry)
											
											'move to the next entery
											rsMove.MoveNext
										loop
										
										rsMove.Close 
										
										strsql = "SELECT tblnews.* FROM tblnews ORDER BY newsiD desc;"
										
										rsMove.Open strsql, adocon
										
										'create the new entery to let the user know they have been moved
										
										Dim lngNewsID
										
										'decide the new news id
										if rsMove.EOF then
												lngnewsid = 1
											else
												lngnewsid = clng(rsmove("newsid"))
												lngnewsid = lngnewsid + 1
										end if
										
										'apply the formatting to the news message
										strNewsMessage = replace(strNewsMessage, "[newcontinent]", lngnewcontinent)
										strNewsMessage = replace(strNewsMessage, "[newcountry]", lngnewcountry)
										strNewsMessage = replace(strNewsMessage, "[oldcontinent]", lngoldcontinent)
										strNewsMessage = replace(strNewsMessage, "[oldcountry]", lngoldcountry)
										strNewsMessage = replace(strNewsMessage, "[admin]", strusername)
										strNewsMessage = replace(strNewsMessage, "[username]", strmovename)
										strNewsMessage = replace(strNewsMessage, "[date]", date())
										strNewsMessage = replace(strNewsMessage, "[time]", time())
										
										
										'add the new news entery
										with rsmove
											.AddNew
											.Fields("newsid") = lngnewsid
											.Fields("username") = strmovename
											.Fields("timemade") = now()
											.Fields("message") = strNewsMessage
											.Fields("UnRead") = cbool(true)
											
											'commence the changes
											.Update 
										end with
										
										'the relocaton is compleate
										Response.Redirect("admin.asp?msg=4")
								end if
						end if
						
						
						'close and kill our objects to save memory
						rsMove.Close 
						
						set rsmove = nothing
				end if
				
		end if
end if
										
										
										
										
										
										
								
									
								
								


'code goes above here

end if
%>
<!---#include file="includes\end_check.asp"--->