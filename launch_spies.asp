<%
option explicit
%>
<!---#include file="includes/check.asp"--->

<!---SPELL CHECKED--->
<%

'declair global varibles
Dim dblBaseAccuracy
Dim intMission
Dim lngunits
Dim intToContinent
Dim intToCountry
Dim rsSpies
Dim dblDynAccuracy
Dim rsSave
Dim lngReportID
Dim intCounter
Dim intTickNo
Dim dblmplandratio
Dim lngNumMP
Dim lngamountLand
Dim strTargetName
Dim lngSpiesLost

'generate the accuracy
randomize
dblBaseAccuracy = cdbl((rnd(second(now())) / 2) + 0.5)

'get the user input

intMission = Request.Form("type")
lngUnits = Request.Form("spies")
intToContinent = Request.Form("continent")
intToCountry = Request.Form("country")

if isnumeric(lngunits) = false then
		Response.Redirect("spies.asp?error=7")
	else
		if lngunits < 1 then
				Response.Redirect("spies.asp?error=7")
			else
				if (isnumeric(intTocontinent) = false) OR (isnumeric(inttocountry) = false) then
						Response.Redirect("spies.asp?error=8")
					else
						if (intToContinent < 1) OR (inttocountry < 1) then
								Response.Redirect("spies.asp?error=8")
							else
								'work out the mission type
								
								'generate the objects that all the missions will use
								set rsSpies = server.CreateObject("ADODB.Recordset")
								
								strsql = "SELECT tickno FROM tblticks;"
								
								rsSpies.open strsql, adocon
								
								intTickNo = cint(rsSpies("tickno"))
								
								rsSpies.Close 
								
								rsSpies.LockType = 3
								
								rsSpies.CursorType = 2
								
								select case cint(intMission)
									case 1
										Dim strUnitArray(18)
										
										strsql = "SELECT numspies FROM tblaccounts WHERE username = '" & strusername & "';"
										
										rsSpies.Open strsql,adocon
										
										if clng(rsspies("numspies")) < clng(lngunits) then
												Response.Redirect("spies.asp?error=13")
											else
												rsSpies.close
										
												strUnitArray(0) = "Infantry"
												strUnitArray(1) = "Comandoes"
												strUnitArray(2) = "Jeeps"
												strUnitArray(3) = "Matilda1"
												strUnitArray(4) = "Matilda2"
												strUnitArray(5) = "Sherman"
												strUnitArray(6) = "Spitfire1"
												strUnitArray(7) = "Spitfire2"
												strUnitArray(8) = "Spitfire5"
												strUnitArray(9) = "Hurricane"
												strUnitArray(10) = "Halifax"
												strUnitArray(11) = "Wellington"
												strUnitArray(12) = "Lancaster"
												strUnitArray(13) = "Landingcraft"
												strUnitArray(14) = "Landingship"
												strUnitArray(15) = "frigates"
												strUnitArray(16) = "battleships"
												strUnitArray(17) = "carrier"
												'the user wants to spy on the targets force that he is sending to attack them
												strsql = "SELECT username, amountland, nummp FROM tblaccounts WHERE continent = " & inttocontinent & " AND countryid = " & inttocountry
												
												rsSpies.Open strsql, adocon
												
												lngnummp = clng(rsspies("nummp"))
												lngamountland = clng(rsspies("amountland"))
												strTargetName = rsspies("username")
												
												dblMPLandRatio = cdbl(clng(rsspies("nummp")) / clng(rsspies("amountland")))
												
												rsSpies.Close 
												
												strsql = "SELECT amountland, numspies FROM tblaccounts WHERE username = '" & strusername & "';"
												
												rsSpies.Open strsql, adocon
												
												if cdbl((clng(lngunits)/clng(rsspies("amountland")))*5) < dblMPLandratio then
														
														if cdbl((clng(lngunits)/clng(rsspies("amountland")))*20) < dblMPLandratio then
																lngSpiesLost = lngUnits
															else
																randomize
																lngSpiesLost = clng(lngUnits * rnd(second(now())))
														end if
														
														rsSpies.Fields("numspies") = clng(rsspies("numspies")) - lngSpiesLost
														
														rsSpies.Update 
														
														rsSpies.Close 
														
														strsql = "SELECT * FROM tblnews ORDER by NEWSID dESC;"
														
														rsSpies.Open strsql, adocon
														
														if rsSpies.EOF then
																lngreportid = 1
															else
																lngreportid = clng(rsspies("newsid")) + 1
														end if
														
														with rsSpies
															.AddNew 
															.Fields("newsid") = lngreportid
															.Fields("username") = strusername
															.Fields("timemade") = now()
															.Fields("message") = "We attempted a spy operation on " & inttocontinent & ":" & inttocountry & " and sadly our spies failed, during the operation " & lngspieslost & " spies were captured and have since been executed.  "
															.Fields("unread") = 1
														
															.Update
														end with
														
														Response.Redirect("spies.asp?error=10")
													else
														
														rsSpies.Close 
												
														strsql = "SELECT eta, infantry, comandoes, jeeps, matilda1, matilda2, sherman, spitfire1, spitfire2, spitfire5, hurricane, lancaster, wellington, halifax, landingcraft, landingship, frigates, battleships, carrier FROM tblmovements WHERE status = 1 AND fromcountry = " & inttocountry & " AND fromcontinent = " & inttocontinent & " AND tocountry = " & lngcountryno & " AND tocontinent = " & lngcountryno
												
														rsSpies.Open strsql, adocon
													
														if rsSpies.EOF then
																Response.Redirect("spies.asp?error=9")
															else
														
																strsql = "SELECT * FROM TblIntelegenceReports ORDER by reportid DESC;"
														
																set rssave = server.CreateObject("ADODB.Recordset")
														
																rsSave.LockType = 3
														
																rsSave.CursorType = 2
														
																rsSave.Open strsql, adocon
															
																if rsSave.EOF then
																		lngreportid = 1
																	else
																		lngreportid = clng(rssave("reportid")) + 1
																end if
														
																do while not rsSpies.EOF 
																	with rssave
																		.AddNew 
																		.Fields("username") = strusername
																		.Fields("accuracy") = formatnumber(cdbl(dblbaseaccuracy*100),2)
																		for intCounter = 0 to 17
																			randomize
																			dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																		
																			dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																		
																			randomize
																			Response.Write(strUnitArray(intCounter))
																			if rnd() > 0.5 then
																					.Fields(strUnitArray(intCounter)) = clng(clng(rsspies(strUnitArray(intCounter))) * (1/dbldynaccuracy))
																				else
																					.Fields(strUnitArray(intCounter)) = clng(clng(rsspies(strUnitArray(intCounter))) * dbldynaccuracy)
																			end if
																		next
																
																		.Fields("reportid") = lngReportID
																		.Fields("countryid") = inttocountry
																		.Fields("continentid") = inttocontinent
																		.Fields("tickreport") = intTickNo
																		
																		.Update
																		
																	end with
																	
																	rsSpies.MoveNext 
																loop
																
																rsSave.Close 
																
																set rssave = nothing
																
																rsSpies.Close 
																	
																strsql = "SELECT tblnews.* FROM tblnews ORDER BY newsID DESC;"
																
																rsSpies.Open strsql, adocon
																		
																if rsSpies.EOF then
																		lngReportid = 1
																	else
																		lngReportid = clng(rsspies("newsid")) + 1
																end if
																
																if clng((lngnummp-lngUnits)/lngamountland) < 0 then
																		'no spies lost
																		
																		with rsspies 
																			.AddNew 
																			.Fields("newsid") = lngreportid
																			.Fields("username") = strusername
																			.Fields("timemade") = now()
																			.Fields("Message") = "Our spy operation was a success. No spies were captured and all have now returned home."
																			.Fields("unread") = 1
																			
																			.Update
																		end with
																	
																	else
																		'clng((lngnummp-lngUnits)/lngamountland) spies lost
																
																		with rsspies 
																			.AddNew 
																			.Fields("newsid") = lngreportid
																			.Fields("username") = strusername
																			.Fields("timemade") = now()
																			.Fields("message") = "Our spies successfully retrieved the requested information, however " & clng((lngnummp-lngUnits)/lngamountland) & " spies were caught and executed."
																			.Fields("unread") = 1
																			
																			.Update
																	
																			.AddNew 
																			.Fields("newsid") = lngreportid + 1
																			.Fields("username") = strTargetName
																			.Fields("timemade") = now()
																			.Fields("message") = "We have noticed spies giving information away to " & lngcontinentno & ":" & lngcountryno & ". We have caught and executed " & clng((lngnummp-lngUnits)/lngamountland) & " spies."
																			.Fields("unread") = 1
																			
																			.Update
																		end with
																
																		rsSpies.Close 
																	
																		strsql = "SELECT numspies FROM tblaccounts WHERE username = '" & strusername & "';"
																
																		rsSpies.Open strsql, adocon
																
																		rsSpies.Fields("numspies") = clng(rsspies("numspies")) - clng((lngnummp-lngUnits)/lngamountland)
																
																		rsSpies.Update 
																end if
																Response.Redirect("spies.asp?error=11")
														end if
												end if
										end if								
														
												
									case 2
										'the user wants to spy on what units the target has at home
										
										
										strsql = "SELECT numspies FROM tblaccounts WHERE username = '" & strusername & "';"
										
										rsSpies.Open strsql,adocon
										
										if clng(rsspies("numspies")) < clng(lngunits) then
												Response.Redirect("spies.asp?error=13")
											else
												rsSpies.close
										
												'the user wants to spy on the targets force that he is sending to attack them
												strsql = "SELECT username, amountland, nummp FROM tblaccounts WHERE continent = " & inttocontinent & " AND countryid = " & inttocountry
												
												rsSpies.Open strsql, adocon
												if rsSpies.EOF then
														Response.Redirect("spies.asp")
													else
												lngnummp = clng(rsspies("nummp"))
												lngamountland = clng(rsspies("amountland"))
												strTargetName = rsspies("username")
												
												dblMPLandRatio = cdbl(clng(rsspies("nummp")) / clng(rsspies("amountland")))
												
												rsSpies.Close 
												
												strsql = "SELECT amountland, numspies FROM tblaccounts WHERE username = '" & strusername & "';"
												
												rsSpies.Open strsql, adocon
												
												if cdbl((clng(lngunits)/clng(rsspies("amountland")))*5) < dblMPLandratio then
														
														if cdbl((clng(lngunits)/clng(rsspies("amountland")))*20) < dblMPLandratio then
																lngSpiesLost = lngUnits
															else
																randomize
																lngSpiesLost = clng(lngUnits * rnd(second(now())))
														end if
														
														rsSpies.Fields("numspies") = clng(rsspies("numspies")) - lngSpiesLost
														
														rsSpies.Update 
														
														rsSpies.Close 
														
														strsql = "SELECT * FROM tblnews ORDER by NEWSID dESC;"
														
														rsSpies.Open strsql, adocon
														
														if rsSpies.EOF then
																lngreportid = 1
															else
																lngreportid = clng(rsspies("newsid")) + 1
														end if
														
														with rsSpies
															.AddNew 
															.Fields("newsid") = lngreportid
															.Fields("username") = strusername
															.Fields("timemade") = now()
															.Fields("message") = "We attempted a spy operation on " & inttocontinent & ":" & inttocountry & " and our spies failed. During the operation, " & lngspieslost & " spies were captured and have been executed."
															.Fields("unread") = 1
														
															.Update
														end with
														
														Response.Redirect("spies.asp?error=10")
													else
														
														rsSpies.Close 
														
														strsql = "SELECT numinfantry, numcommandos, numjeeps, nummat1, nummat2, numsherman, numspit1, numspit2, numspit5, numhurr, numhalifax, numwellington, numlancaster, numlandingcraft, numlandingship, numfrigates, numbattleship, numcarrier, numaaguns, numpilboxes, numturrets, numballoons, numwatermines, numspies, nummp FROM tblaccounts WHERE continent = " & inttocontinent & " AND countryid = " & inttocountry
														
														rsSpies.Open strsql, adocon
														
														if rsSpies.EOF then
																Response.Redirect("spies.asp?error=9")
															else
														
																strsql = "SELECT * FROM TblIntelegenceReports ORDER by reportid DESC;"
														
																set rssave = server.CreateObject("ADODB.Recordset")
														
																rsSave.LockType = 3
														
																rsSave.CursorType = 2
														
																rsSave.Open strsql, adocon
															
																if rsSave.EOF then
																		lngreportid = 1
																	else
																		lngreportid = clng(rssave("reportid")) + 1
																end if
													
																with rssave
																	.AddNew 
																	.Fields("username") = strusername
																	.Fields("accuracy") = formatnumber(cdbl(dblbaseaccuracy*100),2)
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																
																	if rnd() > 0.5 then
																			.Fields("infantry") = clng(clng(rsspies("numinfantry")) * (1/dbldynaccuracy))
																		else
																			.Fields("infantry") = clng(clng(rsspies("numinfantry")) * dbldynaccuracy)
																	end if
																	
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("comandoes") = clng(clng(rsspies("numcommandos")) * (1/dbldynaccuracy))
																		else
																			.Fields("comandoes") = clng(clng(rsspies("numcommandos")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("jeeps") = clng(clng(rsspies("numjeeps")) * (1/dbldynaccuracy))
																		else
																			.Fields("jeeps") = clng(clng(rsspies("numjeeps")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("matilda1") = clng(clng(rsspies("nummat1")) * (1/dbldynaccuracy))
																		else
																			.Fields("matilda1") = clng(clng(rsspies("nummat1")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("matilda2") = clng(clng(rsspies("nummat2")) * (1/dbldynaccuracy))
																		else
																			.Fields("matilda2") = clng(clng(rsspies("nummat2")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("sherman") = clng(clng(rsspies("numsherman")) * (1/dbldynaccuracy))
																		else
																			.Fields("sherman") = clng(clng(rsspies("numsherman")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("spitfire1") = clng(clng(rsspies("numspit1")) * (1/dbldynaccuracy))
																		else
																			.Fields("spitfire1") = clng(clng(rsspies("numspit1")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("spitfire2") = clng(clng(rsspies("numspit2")) * (1/dbldynaccuracy))
																		else
																			.Fields("spitfire2") = clng(clng(rsspies("numspit2")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("spitfire5") = clng(clng(rsspies("numspit5")) * (1/dbldynaccuracy))
																		else
																			.Fields("spitfire5") = clng(clng(rsspies("numspit5")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("hurricane") = clng(clng(rsspies("numhurr")) * (1/dbldynaccuracy))
																		else
																			.Fields("hurricane") = clng(clng(rsspies("numhurr")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("halifax") = clng(clng(rsspies("numhalifax")) * (1/dbldynaccuracy))
																		else
																			.Fields("halifax") = clng(clng(rsspies("numhalifax")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("wellington") = clng(clng(rsspies("numwellington")) * (1/dbldynaccuracy))
																		else
																			.Fields("wellington") = clng(clng(rsspies("numwellington")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("lancaster") = clng(clng(rsspies("numlancaster")) * (1/dbldynaccuracy))
																		else
																			.Fields("lancaster") = clng(clng(rsspies("numlancaster")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("landingcraft") = clng(clng(rsspies("numlandingcraft")) * (1/dbldynaccuracy))
																		else
																			.Fields("landingcraft") = clng(clng(rsspies("numlandingcraft")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("landingship") = clng(clng(rsspies("numlandingship")) * (1/dbldynaccuracy))
																		else
																			.Fields("landingship") = clng(clng(rsspies("numlandingship")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("frigates") = clng(clng(rsspies("numfrigates")) * (1/dbldynaccuracy))
																		else
																			.Fields("frigates") = clng(clng(rsspies("numfrigates")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("battleships") = clng(clng(rsspies("numbattleship")) * (1/dbldynaccuracy))
																		else
																			.Fields("battleships") = clng(clng(rsspies("numbattleship")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("carrier") = clng(clng(rsspies("numcarrier")) * (1/dbldynaccuracy))
																		else
																			.Fields("carrier") = clng(clng(rsspies("numcarrier")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("aaguns") = clng(clng(rsspies("numaaguns")) * (1/dbldynaccuracy))
																		else
																			.Fields("aaguns") = clng(clng(rsspies("numaaguns")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("pilboxes") = clng(clng(rsspies("numpilboxes")) * (1/dbldynaccuracy))
																		else
																			.Fields("pilboxes") = clng(clng(rsspies("numpilboxes")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("turrets") = clng(clng(rsspies("numturrets")) * (1/dbldynaccuracy))
																		else
																			.Fields("turrets") = clng(clng(rsspies("numturrets")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																		
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("balloons") = clng(clng(rsspies("numballoons")) * (1/dbldynaccuracy))
																		else
																			.Fields("balloons") = clng(clng(rsspies("numballoons")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("mines") = clng(clng(rsspies("numwatermines")) * (1/dbldynaccuracy))
																		else
																			.Fields("mines") = clng(clng(rsspies("numwatermines")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("spies") = clng(clng(rsspies("numspies")) * (1/dbldynaccuracy))
																		else
																			.Fields("spies") = clng(clng(rsspies("numspies")) * dbldynaccuracy)
																	end if
																	
																	randomize
																	dblDynAccuracy = cdbl((rnd(minute(now())) - 0.5)/6)
																	
																	dbldynAccuracy  = dblbaseaccuracy + dbldynaccuracy
																	
																	randomize
																	if rnd() > 0.5 then
																			.Fields("mp") = clng(clng(rsspies("nummp")) * (1/dbldynaccuracy))
																		else
																			.Fields("mp") = clng(clng(rsspies("nummp")) * dbldynaccuracy)
																	end if
																	.Fields("reportid") = lngReportID
																	.Fields("countryid") = inttocountry
																	.Fields("continentid") = inttocontinent
																	.Fields("tickreport") = intTickNo
																	
																	.Update
																	
																end with
																
																																
																rsSave.Close 
																
																set rssave = nothing
																
																rsSpies.Close 
																	
																strsql = "SELECT tblnews.* FROM tblnews ORDER BY newsID DESC;"
																
																rsSpies.Open strsql, adocon
																		
																if rsSpies.EOF then
																		lngReportid = 1
																	else
																		lngReportid = clng(rsspies("newsid")) + 1
																end if
																
																if clng((lngnummp-lngUnits)/lngamountland) < 0 then
																		'no spies lost
																		
																		with rsspies 
																			.AddNew 
																			.Fields("newsid") = lngreportid
																			.Fields("username") = strusername
																			.Fields("timemade") = now()
																			.Fields("Message") = "Our spy operation was a success. No spies were captured and all have now returned home."
																			.Fields("unread") = 1
																			
																			.Update
																		end with
																	
																	else
																		'clng((lngnummp-lngUnits)/lngamountland) spies lost
																
																		with rsspies 
																			.AddNew 
																			.Fields("newsid") = lngreportid
																			.Fields("username") = strusername
																			.Fields("timemade") = now()
																			.Fields("message") = "Our spies successfully retrieved the requested information, however " & clng((lngnummp-lngUnits)/lngamountland) & " spies were caught and executed."
																			.Fields("unread") = 1
																			
																			.Update
																	
																			.AddNew 
																			.Fields("newsid") = lngreportid + 1
																			.Fields("username") = strTargetName
																			.Fields("timemade") = now()
																			.Fields("message") = "We have noticed spies giving information away to " & lngcontinentno & ":" & lngcountryno & ". We have caught and executed " & clng((lngnummp-lngUnits)/lngamountland) & " spies."
																			.Fields("unread") = 1
																			
																			.Update
																		end with
																
																		rsSpies.Close 
																	
																		strsql = "SELECT numspies FROM tblaccounts WHERE username = '" & strusername & "';"
																
																		rsSpies.Open strsql, adocon
																
																		rsSpies.Fields("numspies") = clng(rsspies("numspies")) - clng((lngnummp-lngUnits)/lngamountland)
																
																		rsSpies.Update 
																end if
																Response.Redirect("spies.asp?error=11")
														end if
												end if
											end if
										end if
														
									case 3
										'the user wants to send in spies to take out the users refineries
										
										strsql = "SELECT numspies FROM tblaccounts WHERE username = '" & strusername & "';"
										
										rsSpies.Open strsql,adocon
										
										if clng(rsspies("numspies")) < clng(lngunits) then
												Response.Redirect("spies.asp?error=13")
											else
												rsSpies.close
										
												'the user wants to spy on the targets force that he is sending to attack them
												strsql = "SELECT username, amountland, nummp FROM tblaccounts WHERE continent = " & inttocontinent & " AND countryid = " & inttocountry
												
												rsSpies.Open strsql, adocon
												
												lngnummp = clng(rsspies("nummp"))
												lngamountland = clng(rsspies("amountland"))
												strTargetName = rsspies("username")
												if clng(rsspies("amountland")) = 0 then
														dblmplandratio = clng(rsspies("nummp")) * 10
													else
														dblMPLandRatio = cdbl(cdbl(rsspies("nummp")) / clng(rsspies("amountland")))
												end if
												
												rsSpies.Close 
												
												strsql = "SELECT amountland, numspies FROM tblaccounts WHERE username = '" & strusername & "';"
												
												rsSpies.Open strsql, adocon
												
												if cdbl((clng(lngunits)/clng(rsspies("amountland")))*2) < dblMPLandratio then
												
														if cdbl((clng(lngunits)/clng(rsspies("amountland")))*20) < dblMPLandratio then
																lngSpiesLost = lngUnits
															else
																randomize
																lngSpiesLost = clng(lngUnits * rnd(second(now())))
														end if
														
														rsSpies.Fields("numspies") = clng(rsspies("numspies")) - lngSpiesLost
														
														rsSpies.Update 
														
														rsSpies.Close 
														
														strsql = "SELECT * FROM tblnews ORDER by NEWSID dESC;"
														
														rsSpies.Open strsql, adocon
														
														if rsSpies.EOF then
																lngreportid = 1
															else
																lngreportid = clng(rsspies("newsid")) + 1
														end if
														
														with rsSpies
															.AddNew 
															.Fields("newsid") = lngreportid
															.Fields("username") = strusername
															.Fields("timemade") = now()
															.Fields("message") = "We attempted a spy operation on " & inttocontinent & ":" & inttocountry & " and our spies failed. During the operation, " & lngspieslost & " spies were captured and have been executed."
															.Fields("unread") = 1
														
															.Update
														end with
														
														Response.Redirect("spies.asp?error=10")
													else
														rsSpies.Close 
														'blow up a refinery
														'save the news
														'go home
														
														strsql = "SELECT numrefineries FROM tblaccounts WHERE username = '" & strTargetName & "';"
														
														rsSpies.Open strsql,adocon 
														
														if clng(rsspies("numrefineries")) < 1 then
																Response.Redirect("spies.asp?error=14")
															else
															
														
																rsSpies.Fields("numrefineries") = clng(rsspies("numrefineries")) - 1
																
																rsSpies.Update
																
																rsSpies.Close 
																
																strsql = "SELECT * FROM tblnews ORDER BY newsid DESC;"
																
																
																rsSpies.Open strsql, adocon
																
																if rsSpies.EOF then
																		lngreportid = 1
																	else
																		lngreportid = clng(rsspies("newsid")) + 1
																end if
																
																with rsSpies
																	.AddNew 
																	.Fields("newsid") = lngreportid
																	.Fields("username") = strusername
																	.Fields("timemade") = now()
																	if clng(((lngnummp-lngUnits)/lngamountland)*5) >= 1 then
																			.Fields("message") = "Our spies have targeted a power surge attack on " & inttocontinent & ":" & inttocountry & " and have blown up a refinery. " & clng(((lngnummp-lngUnits)/lngamountland)*5) & " spies were killed in the attack."
																		else
																			.Fields("message") = "Our spies have targeted a power surge attack on " & inttocontinent & ":" & inttocountry & " and have blown up a refinery. No spies were killed in the attack."
																	end if
																	.Fields("unread") = 1
																	
																	.Update
																	
																	.AddNew 
																	.Fields("newsid") = lngreportid + 1
																	.Fields("username") = strtargetname
																	.Fields("timemade") = now()
																	.Fields("message") = "One of our refineries has been subjected to a power surge and has exploded."
																	.Fields("unread") = 1
																	.update
																	
																end with
																
																if clng(((lngnummp-lngUnits)/lngamountland)*5) >= 1 then
																		rsSpies.Close 
																		
																		strsql = "SELECT numspies FROM tblaccounts WHERE username = '" & strusername & "';"
																		
																		rsSpies.Open strsql, adocon
																		
																		if clng(((lngnummp-lngUnits)/lngamountland)*5) > clng(rsspies("numspies")) then
																				rsSpies.Fields("numspies") = 0
																			else
																				rsSpies.Fields("numspies") = clng(rsspies("numspies")) - clng(((lngnummp-lngUnits)/lngamountland)*5)
																		end if
																		
																		rsSpies.Update
																end if 
																
																Response.Redirect("spies.asp?error=11")
														end if
														
														
												end if
										end if
												
												
									case 4
										'the user wants to take out one of the enemies bomber planes
										
										
										strsql = "SELECT numspies FROM tblaccounts WHERE username = '" & strusername & "';"
										
										rsSpies.Open strsql,adocon
										
										if clng(rsspies("numspies")) < clng(lngunits) then
												Response.Redirect("spies.asp?error=13")
											else
												rsSpies.close
										
												'the user wants to spy on the targets force that he is sending to attack them
												strsql = "SELECT username, amountland, nummp FROM tblaccounts WHERE continent = " & inttocontinent & " AND countryid = " & inttocountry
												
												rsSpies.Open strsql, adocon
												
												lngnummp = clng(rsspies("nummp"))
												lngamountland = clng(rsspies("amountland"))
												strTargetName = rsspies("username")
												
												dblMPLandRatio = cdbl(clng(rsspies("nummp")) / clng(rsspies("amountland")))
												
												rsSpies.Close 
												
												strsql = "SELECT amountland, numspies FROM tblaccounts WHERE username = '" & strusername & "';"
												
												rsSpies.Open strsql, adocon
												
												if cdbl((clng(lngunits)/clng(rsspies("amountland")))*2) < dblMPLandratio then
												
														if cdbl((clng(lngunits)/clng(rsspies("amountland")))*20) < dblMPLandratio then
																lngSpiesLost = lngUnits
															else
																randomize
																lngSpiesLost = clng(lngUnits * rnd(second(now())))
														end if
														
														rsSpies.Fields("numspies") = clng(rsspies("numspies")) - lngSpiesLost
														
														rsSpies.Update 
														
														rsSpies.Close 
														
														strsql = "SELECT * FROM tblnews ORDER by NEWSID dESC;"
														
														rsSpies.Open strsql, adocon
														
														if rsSpies.EOF then
																lngreportid = 1
															else
																lngreportid = clng(rsspies("newsid")) + 1
														end if
														
														with rsSpies
															.AddNew 
															.Fields("newsid") = lngreportid
															.Fields("username") = strusername
															.Fields("timemade") = now()
															.Fields("message") = "We attempted a spy operation on " & inttocontinent & ":" & inttocountry & " and our spies failed. During the operation, " & lngspieslost & " spies were captured and have been executed.  "
															.Fields("unread") = 1
														
															.Update
														end with
														
														Response.Redirect("spies.asp?error=10")
													else
														rsSpies.Close 
														'blow up a refinery
														'save the news
														'go home
														
														strsql = "SELECT numhalifax, numwellington, numlancaster FROM tblaccounts WHERE username = '" & strTargetName & "';"
														
														rsSpies.Open strsql,adocon 
														
														if clng(clng(rsspies("numlancaster"))+clng(rsspies("numwellington"))+clng(rsspies("numhalifax"))) < 1 then
																Response.Redirect("spies.asp?error=15")
															else
															
														
																if clng(rsspies("numlancaster")) > 0 then
																		rsSpies.Fields("numlancaster") = clng(rsspies("numlancaster")) - 1
																	else
																		if clng(rsspies("numwellington")) > 0 then
																				rsSpies.Fields("numwellington") = clng(rsspies("numwellington")) - 1
																			else
																				rsSpies.Fields("numhalifax") = clng(rsspies("numhalifax")) - 1
																		end if
																end if
																										
																rsSpies.Update
																
																rsSpies.Close 
																
																strsql = "SELECT * FROM tblnews ORDER BY newsid DESC;"
																
																
																rsSpies.Open strsql, adocon
																
																if rsSpies.EOF then
																		lngreportid = 1
																	else
																		lngreportid = clng(rsspies("newsid")) + 1
																end if
																
																with rsSpies
																	.AddNew 
																	.Fields("newsid") = lngreportid
																	.Fields("username") = strusername
																	.Fields("timemade") = now()
																	if clng(((lngnummp-lngUnits)/lngamountland)*5) >= 1 then
																			.Fields("message") = "Our spies have arrived in " & inttocontinent & ":" & inttocountry & " and rendered a bomber useless. " & clng(((lngnummp-lngUnits)/lngamountland)*5) & " spies were killed in the attack."
																		else
																			.Fields("message") = "Our spies have arrived in " & inttocontinent & ":" & inttocountry & " and rendered a bomber useless. No spies were killed in the attack."
																	end if
																	.Fields("unread") = 1
																	
																	.Update
																	
																	.AddNew 
																	.Fields("newsid") = lngreportid + 1
																	.Fields("username") = strtargetname
																	.Fields("timemade") = now()
																	.Fields("message") = "During an inspection, one of our bombers was found to be in poor condition. The damage was such the bomber has been scapped."
																	.Fields("unread") = 1
																	.update
																	
																end with
																
																if clng(((lngnummp-lngUnits)/lngamountland)*5) >= 1 then
																		rsSpies.Close 
																		
																		strsql = "SELECT numspies FROM tblaccounts WHERE username = '" & strusername & "';"
																		
																		rsSpies.Open strsql, adocon
																		
																		if clng(((lngnummp-lngUnits)/lngamountland)*5) > clng(rsspies("numspies")) then
																				rsSpies.Fields("numspies") = 0
																			else
																				rsSpies.Fields("numspies") = clng(rsspies("numspies")) - clng(((lngnummp-lngUnits)/lngamountland)*5)
																		end if
																		
																		rsSpies.Update
																end if 
																
																Response.Redirect("spies.asp?error=11")
														end if
														
														
												end if
										end if
								end select
								
								rsSpies.Close
								
								set rsspies = nothing
						end if
				end if
		end if
end if
										

												
												
%>
<!---#include file="includes/end_check.asp"--->