<%
'SPELL CHECKED

option explicit
%>
<!---#include file="includes\check.asp"--->
<%
Dim rsunits			'recordset to peform the work
Dim intCurMissions
Dim lngInfantry
Dim lngCommandoes
Dim lngJeeps
Dim lngMatilda1
Dim lngMatilda2
Dim lngSherman
Dim lngSpitfire1
Dim lngSpitfire2
Dim lngSpitfire5
Dim lngHurricane
Dim lngHalifax
Dim lngWellington
Dim lngLancaster
Dim lngFrigates
Dim lngBattleships
Dim lngCarriers
Dim intToContinent		'the continent the force is going to
Dim intToCountry		'the country the force is going to
Dim intMissionType		'1=attack, 2=defend
dim intTravelTime		'the time needed to get to the target
Dim lngEnemyScore		'used to stop newbie bashing
Dim lngLandingCraft
Dim lngLandingShip
Dim lngOrderNo			'id number for the movement
Dim strTargetName		'username of the target for news
Dim lngNewsID			'id number for the news
Dim numattckers

'get the user input
lngInfantry = Request.Form("infantry")
lngCommandoes = Request.Form("commandoes")
lngJeeps = Request.Form("jeeps")
lngMatilda1 = Request.Form("mat1")
lngMatilda2 = Request.Form("mat2")
lngSherman = Request.Form("sherman")
lngspitfire1 = Request.Form("spitfire1")
lngspitfire2 = Request.Form("spitfire2")
lngspitfire5 = Request.Form("spitfire5")
lngHurricane = Request.Form("hurricane")
lnghalifax = Request.Form("halifax")
lngwellington = Request.Form("wellington")
lngLancaster = Request.Form("lancaster")
lngfrigates = Request.Form("frigates")
lngbattleships = Request.Form("battleships")
lngcarriers = Request.Form("carrier")
inttocontinent = Request.Form("continent")
inttocountry = Request.Form("country")
intmissiontype = Request.Form("type")

'validate the user input
if isnumeric(lngInfantry) = false then
		lngInfantry = 0
	else
		if lngInfantry < 0 then
				lngInfantry = 0
		end if
end if

if isnumeric(lngCommandoes) = false then
		lngCommandoes = 0
	else
		if lngCommandoes < 0 then
				lngCommandoes = 0
		end if
end if

if isnumeric(lngJeeps) = false then
		lngJeeps = 0
	else
		if lngJeeps < 0 then
				lngJeeps = 0
		end if
end if

if isnumeric(lngMatilda1) = false then
		lngMatilda1 = 0
	else
		if lngMatilda1 < 0 then
				lngMatilda1 = 0
		end if
end if

if isnumeric(lngMatilda2) = false then
		lngMatilda2 = 0
	else
		if lngMatilda2 < 0 then
				lngMatilda2 = 0
		end if
end if

if isnumeric(lngspitfire1) = false then
		lngspitfire1 = 0
	else
		if lngSpitfire1 < 0 then
				lngspitfire1 = 0
		end if
end if

if isnumeric(lngSpitfire2) = false then
		lngSpitfire2 = 0
	else
		if lngSpitfire2 < 0 then
				lngSpitfire2 = 0
		end if
end if

if isnumeric(lngSpitfire5) = false then
		lngSpitfire5 = 0
	else
		if lngSpitfire5 < 0 then
				lngspitfire5 = 0
		end if
end if

if isnumeric(lngHurricane) = false then
		lngHurricane = 0
	else
		if lngHurricane < 0 then
				lngHurricane = 0
		end if
end if

if isnumeric(lngHalifax) = false then
		lngHalifax = 0
	else
		if lngHalifax < 0 then
				lngHalifax = 0
		end if
end if

if isnumeric(lngWellington) = false then
		lngWellington = 0
	else
		if lngwellington < 0 then
				lngWellington = 0
		end if
end if

if isnumeric(lngLancaster) = false then
		lngLancaster = 0
	else
		if lngLancaster < 0 then
				lngLancaster = 0
		end if
end if

if isnumeric(lngfrigates) = false then
		lngfrigates = 0
	else
		if lngFrigates < 0 then
				lngFrigates = 0
		end if
end if

if isnumeric(lngBattleships) = false then
		lngBattleships = 0
	else
		if lngBattleships < 0 then
				lngBattleships = 0
		end if
end if

if isnumeric(lngCarriers) = false then
		lngCarriers = 0
	else
		if lngCarriers < 0 then
				lngCarriers = 0
		end if
end if

if blnTickRunning = true then
		Response.Redirect("mil_opps.asp?error=9")
	else
if isnumeric(intMissionType) = false then
		Response.Redirect("mil_opps.asp")
		'the user has made their own form
	else
		if (intMissionType <> 1) AND (intMissiontype <> 2) then
				Response.Redirect("mil_opps.asp")
				'the mission type is invalid, user has made their own form
			else
				if (isnumeric(inttocontinent) = false) OR (isnumeric(inttocountry) = false) then
						Response.Redirect("mil_opps.asp?error=1")
						'user has not entered a numeic country location
					else
						if (lngInfantry + lngCommandoes + lngjeeps + lngmatilda1 + lngmatilda2 + lngspitfire1 + lngspitfire2 + lngspitfire5 + lngsherman + lnghurricane + lngwellington + lnghalifax + lnglancaster + lngfrigates + lngbattleships + lngcarriers) = 0 then
								Response.Redirect("mil_opps.asp?error=2")
								'user has tried to send 0 units
							else
								'created the objects needed to open the database for further validation
								set Rsunits = server.CreateObject("ADODB.Recordset")
								
								strsql = "SELECT score, username FROM tblaccounts WHERE continent = " & inttocontinent & " AND countryid = " & inttocountry
								
								rsunits.Open strsql, adocon
								
								if rsunits.EOF then
										'the target does not exist
										Response.Redirect("mil_opps.asp?error=3")
									else
										'save score and username for later use
										lngEnemyScore = rsunits("score")
										strTargetName = rsunits("username")
										rsunits.Close 
										
										
										rsunits.LockType = 3
										rsunits.CursorType = 2
										
										
										strsql = "SELECT numinfantry, numcommandos, numjeeps, nummat1, nummat2, numsherman, numspit1, numspit2, numspit5, numhurr, numlancaster, numwellington, numhalifax, numlandingcraft, numlandingship, numfrigates, numbattleship, numcarrier, amountland, score FROM tblaccounts WHERE username = '" & strusername & "';"
										
										rsunits.Open strsql, adocon
										Dim blnAllowed
										'check to make sure the target is big enough if attacking
										if clng(clng(rsunits("score"))/3) > clng(lngEnemyScore) then
												if intmissiontype = 1 then
														blnAllowed = false
													else
														blnAllowed = true
												end if
											else
												blnAllowed = true
										end if
										
										if blnAllowed = false then
												'target is too small
												Response.Redirect("mil_opps.asp?error=4")
											else
												'the attack is allowed, make sure the user has enough units for the attack
												
												if clng(rsunits("numinfantry")) < clng(lnginfantry) then
														lngInfantry = clng(rsunits("numinfantry"))
												end if
												
												if clng(rsunits("numcommandos")) < clng(lngcommandoes) then
														lngcommandoes = clng(rsunits("numcommandos"))
												end if
												
												if clng(rsunits("numjeeps")) < clng(lngjeeps) then
														lngJeeps = clng(rsunits("numjeeps"))
												end if
												
												if clng(rsunits("nummat1")) < clng(lngmatilda1) then
														lngmatilda1 = clng(rsunits("nummat1"))
												end if
												
												if clng(rsunits("nummat2")) < clng(lngmatilda2) then
														lngmatilda2 = clng(rsunits("nummat2"))
												end if
												
												if clng(rsunits("numsherman")) < clng(lngsherman) then
														lngsherman = clng(rsunits("numsherman"))
												end if
												
												if clng(rsunits("numspit1")) < clng(lngspitfire1) then
														lngspitfire1 = clng(rsnuits("numspit1"))
												end if
												
												if clng(rsunits("numspit2")) < clng(lngspitfire2) then
														lngspitfire2 = clng(rsunits("numspit2"))
												end if
												
												if clng(rsunits("numspit5")) < clng(lngspitfire5) then
														lngspitfire5 = clng(rsunits("numspit5"))
												end if
												
												if clng(rsunits("numhurr")) < clng(lnghurricane) then
														lnghurricane = clng(rsunits("numhurr"))
												end if 
												
												if clng(rsunits("numhalifax")) < clng(lnghalifax) then
														lnghalifax = clng(rsunits("numhalifax"))
												end if
												
												if clng(rsunits("numwellington")) < clng(lngwellington) then
														lngwellington = clng(rsunits("wellington"))
												end if
												
												if clng(rsunits("numlancaster")) < clng(lnglancaster) then
														lnglancaster = clng(rsunits("numlancaster"))
												end if
												
												if clng(rsunits("numfrigates")) < clng(lngfrigates) then
														lngfrigates = clng(rsunits("numfrigates"))
												end if
												
												if clng(rsunits("numbattleship")) < clng(lngbattleships) then
														lngbattleships = clng(rsunits("numbattleships"))
												end if
												
												if clng(rsunits("numcarrier")) = clng(lngcarriers) then
														lngcarriers = clng(rsunits("numcarrier"))
												end if
												
												'check to see if the force is leaving the continent
												if cint(inttocontinent) = cint(lngcontinentno) then
														'no no landing craft required
														lnglandingship = 0
														lnglandingcraft = 0
														
														'set lowest possible travel time
														intTravelTime = 2
														
														if (lnglancaster > 0) OR (lngjeeps > 0) OR (lnghalifax > 0) OR (lngwellington > 0) OR (lngsherman > 0) then
																intTravelTime = 3
														end if
														if (lngmatilda1 > 0) OR (lngmatilda2 > 0) then
																intTravelTime = 4
														end if
														if (lnginfantry > 0) OR (lngcommandoes > 0) then
																intTravelTime = 5
														end if
														if (lngfrigates > 0) OR (lngbattleships > 0) then
																intTravelTime = 6
														end if
														if lngcarriers > 0 then
																inttravelTime = 7
														end if
													else
														'they are leaving their continent
														'set the lowest travel time
														intTravelTime = 4
														if (lnglancaster > 0) OR (lnghalifax > 0) OR (lngwellington > 0) then
																intTraveltime = 6
														end if
														
														'check to see how many landing craft will be needed
														lngLandingCraft = 0
														lngLandingship = 0
														
														lngLandingCraft = clng(((clng(lnginfantry)+clng(lngcommandoes))/100)+0.5)
														lngLandingShip = clng((((clng(lngjeeps)+clng(lngmatilda1)+clng(lngmatilda2)+clng(lngsherman)))/30)+0.5)
														
														if lngLandingCraft > 0 then
																intTravelTime = 8
														end if
														
														if (lngfrigates > 0) OR (lngbattleships > 0) then
																intTravelTime = 9
														end if
														
														if (lnglandingship > 0) then
																intTravelTime = 10
														end if
														
														if lngcarriers > 0 then
																intTravelTime = 12
														end if
												end if
												rsunits.Close 
												strsql = "SELECT status FROM tblmovements WHERE ETA = " & intTravelTime & " AND tocontinent = " & inttocontinent & " AND tocountry = " & inttocountry & " AND status = " & intmissiontype
												
												rsunits.Open strsql, adocon
												
												numattckers = 0
												
												do while not rsunits.EOF 
													numattckers = numattckers + 1
													rsunits.MoveNext 
												loop
												rsunits.Close 
												if numattckers > 23 then
														Response.Redirect("mil_opps.asp?error=")
													else
														strsql = "SELECT numinfantry, numcommandos, numjeeps, nummat1, nummat2, numsherman, numspit1, numspit2, numspit5, numhurr, numlancaster, numwellington, numhalifax, numlandingcraft, numlandingship, numfrigates, numbattleship, numcarrier, amountland FROM tblaccounts WHERE username = '" & strusername & "';"
														
														rsunits.Open strsql, adocon
														'make sure the user has enough landing craft to support the attack
														if (clng(lnglandingcraft) > clng(rsunits("numlandingcraft"))) OR (clng(lnglandingship) > clng(rsunits("numlandingship"))) then
																Response.Redirect("mil_opps.asp?error=5&ship=" & lnglandingship & "&craft=" & lnglandingcraft)
															else
																'double check to make sure the final values are not 0
																if (lngInfantry + lngCommandoes + lngjeeps + lngmatilda1 + lngmatilda2 + lngspitfire1 + lngspitfire2 + lngspitfire5 + lngsherman + lnghurricane + lngwellington + lnghalifax + lnglancaster + lngfrigates + lngbattleships + lngcarriers) = 0 then
																		'yes they are
																		Response.Redirect("mil_opps.asp?error=2")
																	else
																			
																			
																		'no commence the movement
																		'modify the amount of units at home
																		rsunits.Fields("numinfantry") = clng(rsunits("numinfantry")) - lnginfantry
																		rsunits.Fields("numcommandos") = clng(rsunits("numcommandos")) - lngcommandoes
																		rsunits.Fields("numjeeps") = clng(rsunits("numjeeps")) - lngjeeps
																		rsunits.Fields("nummat1") = clng(rsunits("nummat1")) - lngmatilda1
																		rsunits.Fields("nummat2") = clng(rsunits("nummat2")) - lngmatilda2
																		rsunits.Fields("numsherman") = clng(rsunits("numsherman")) - lngsherman
																		rsunits.Fields("numspit1") = clng(rsunits("numspit1")) - lngspitfire1
																		rsunits.Fields("numspit2") = clng(rsunits("numspit2")) - lngspitfire2
																		rsunits.Fields("numspit5") = clng(rsunits("numspit5")) - lngspitfire5
																		rsunits.Fields("numhurr") = clng(rsunits("numhurr")) - lnghurricane
																		rsunits.Fields("numhalifax") = clng(rsunits("numhalifax")) - lnghalifax
																		rsunits.Fields("numwellington") = clng(rsunits("numwellington")) - lngwellington
																		rsunits.Fields("numlancaster") = clng(rsunits("numlancaster")) - lnglancaster
																		rsunits.Fields("numfrigates") = clng(rsunits("numfrigates")) - lngfrigates
																		rsunits.Fields("numbattleship") = clng(rsunits("numbattleship")) - lngbattleships
																		rsunits.Fields("numcarrier") = clng(rsunits("numcarrier")) - lngcarriers
																		rsunits.Fields("numlandingcraft") = clng(rsunits("numlandingcraft")) - lnglandingcraft
																		rsunits.Fields("numlandingship") = clng(rsunits("numlandingship")) - lnglandingship
														
																		rsunits.Update
														
																		rsunits.Close 
												
																		strsql = "SELECT * FROM tblmovements ORDER BY orderno DESC;"
												
																		rsunits.Open strsql, adocon
																			
																		'calculte the new primary key
																		if rsunits.EOF then
																				lngOrderNo = 1
																			else
																				lngOrderNo = clng(rsunits("Orderno")) + 1
																		end if
																		'commence the movement
																		with rsunits
																			.AddNew
																			.Fields("username") = strusername
																			.Fields("ETA") = intTravelTime
																			.Fields("infantry") = lnginfantry
																			.Fields("Comandoes") = lngcommandoes
																			.Fields("jeeps") = lngjeeps
																			.Fields("matilda1") = lngmatilda1
																			.Fields("matilda2") = lngmatilda2
																			.Fields("sherman") = lngsherman
																			.Fields("spitfire1") = lngspitfire1
																			.Fields("spitfire2") = lngspitfire2
																			.Fields("spitfire5") = lngspitfire5
																			.Fields("hurricane") = lnghurricane
																			.Fields("lancaster") = lnglancaster 
																			.Fields("Wellington") = lngwellington
																			.Fields("Halifax") = lnghalifax
																			.Fields("landingcraft") = lnglandingcraft
																			.Fields("landingship") = lnglandingship
																			.Fields("frigates") = lngfrigates
																			.Fields("battleships") = lngbattleships
																			.Fields("carrier") = lngcarriers
																			.Fields("cargoship") = 0
																			.Fields("spies") = 0
																			.Fields("status") = intMissionType
																			.Fields("orderno") = lngorderno
																			.Fields("fromcontinent") = lngcontinentno
																			.Fields("fromcountry") = lngcountryno
																			.Fields("tocontinent") = inttocontinent
																			.Fields("tocountry") = inttocountry
																			.Fields("ticktime") = inttravelTime
																			'save
																			.Update
																		end with
														
																		rsunits.Close 
														
																		strsql = "SELECT tblnews.* FROM tblnews ORDER BY NEWSID DESC;"		
														
																		rsunits.Open strsql, adocon
																		'add news for the two accounts
																		if rsunits.EOF then
																				lngnewsID = 1 
																			else
																				lngnewsID = clng(rsunits("newsid")) + 1
																		end if
														
																		with rsunits
																			.AddNew 
																			.Fields("newsid") = lngnewsid
																			.Fields("username") = strusername
																			.Fields("timemade") = now()
																			if intmissiontype = 1 then
																					.Fields("message") = "We have launched an hostile force to attack " & inttocontinent & ":" & inttocountry & ". The ETA is " & inttraveltime & " ticks."
																				else
																					.Fields("message") = "We have launched a friendly force to defend " & inttocontinent & ":" & inttocountry & ". The ETA is " & inttraveltime & " ticks."
																			end if
																			.Fields("unread") = 1
																				
																			.Update
																				
																			.AddNew 
																			.Fields("newsid") = lngnewsid + 1
																			.Fields("username") = strtargetname
																			.Fields("timemade") = now()
																			if intmissiontype = 1 then
																					.Fields("message") = "We have detected an incoming hostile force from " & lngcontinentno & ":" & lngcountryno & ". The ETA is " & inttraveltime & " ticks."
																				else
																					.Fields("message") = "We have detected an incoming friendly force from " & lngcontinentno & ":" & lngcountryno & ". The ETA is " & inttraveltime & " ticks."
																			end if
																			.Fields("unread") = 1 
																				
																			.Update 
																		end with
																		'mission launched
																		Response.Redirect("mil_opps.asp?error=6")
																end if
														end if
												end if
										end if
								end if
								
								rsunits.Close 
								set rsunits = nothing
						end if
				end if
		end if
end if
end if	

%>
<!---#include file="includes\end_check.asp"--->
