<%
'SPELL CHECKED

option explicit
%>
<!---#include file="includes/check.asp"--->
<%
Dim rsSpies
Dim lngSpies
Dim lngMPs
Dim lngOrderNo
Dim lngFoodCost
Dim lngWoodCost
Dim lngIronCost
Dim lngMoneyCost


'get the user input
lngSpies = Request.Form("spies")
lngMPs = Request.Form("mp")

'validate the user input
if isnumeric(lngSpies) = false then
		lngSpies = 0
	else
		if clng(lngspies) < 0 then
				lngSpies = 0
		end if
end if

if isnumeric(lngMPs) = false then
		lngMPs = 0
	else
		if clng(lngMPs) < 0 then
				lngMPs = 0
		end if
end if
if blnTickRunning = true then
		Response.Redirect("spies.asp?error=16")
	else
'make sure the user is not trying to order nothing
if clng(lngSpies+lngMPs) < 0 then
		Response.Redirect("spies.asp?error=1")
	else
		'calculate the cost
		lngFoodCost = lngMPs * 60
		lngWoodCost = 0
		lngIronCost = lngMPs * 5
		lngMoneyCost = lngMPs * 150
		
		lngfoodCost = lngFoodcost + (lngspies * 100)
		lngWoodCost = lngSpies
		lngIronCost = lngIronCost + (lngspies * 5)
		lngmoneyCost = lngMoneyCost + (lngspies * 400)
		
		'the cost has been calculated
		
		'make sure the user has enough to pay for it
		if lngFoodCost > lngfood then
				Response.Redirect("spies.asp?error=2")
			else
				if clng(lngWoodCost) > clng(lngWood) then
						Response.Redirect("spies.asp?error=3")
					else
						if lngIronCost > lngIron then
								Response.Redirect("spies.asp?error=4")
							else
								if lngmoneycost > lngmoney then
										Response.Redirect("spies.asp?error=5")
									else
										'create the order
										
										'instanciate the objects
										set rsspies = server.CreateObject("ADODB.recordset")
										
										rsSpies.LockType = 3
										
										rsSpies.CursorType = 2
										
										strsql = "SELECT money, wood, food, iron FROM tblaccounts WHERE username = '" & strusername & "';"
										
										rsSpies.Open strsql, adocon
										
										with rsspies
											.Fields("money") = clng(rsspies("money")) - lngmoneycost
											.Fields("wood") = clng(rsspies("wood")) - lngwoodcost
											.Fields("food") = clng(rsspies("food")) - lngfoodcost
											.Fields("iron") = clng(rsspies("iron")) - lngironcost
											
											.Update
										end with
										
										rsSpies.Close 
										
										strsql = "SELECT orderno FROM tblpendingunits ORDER BY orderno DESC;"
										
										rsSpies.Open strsql, adocon
										
										if rsSpies.EOF then
												lngorderno = 1
											else
												lngorderno = clng(rsspies("orderno")) + 1
										end if
										
										rsSpies.Close  
										
										if lngMPs > 0 then
												strsql = "SELECT username, timeleft, orderno, mp FROM tblpendingunits WHERE timeleft = 6 AND username = '" & strusername & "';"
												
												rsSpies.Open strsql, adocon
												
												if rsSpies.EOF then
														with rsSpies
															.addnew
															.fields("username") = strusername
															.fields("timeleft") = 6
															.fields("orderno") = lngorderno
															.fields("mp") = lngmps
															
															.update
														end with
														
														lngorderno = lngorderno + 1
													else
														with rsspies
															.Fields("mp") = clng(rsspies("mp")) + lngmps
															
															.Update
														end with
												end if
												
												rsSpies.Close 
										end if
										
										if lngspies > 0 then
												strsql = "SELECT username, timeleft, orderno, spies FROM tblpendingunits WHERE timeleft = 12 AND username = '" & strusername & "';"
											
												rsSpies.Open strsql, adocon
												
												if rsSpies.EOF then
														with rsspies
															.addnew
															.fields("username") = strusername
															.fields("timeleft") = 12
															.fields("orderno") = lngorderno
															.fields("spies") = lngspies
															
															.update
														end with
													else
														with rsspies 
															.Fields("spies") = clng(rsspies("spies")) + lngspies
															
															.Update
														end with
												end if
												
												rsSpies.Close 
										end if
										
										set rsspies = nothing
										
										Response.Redirect("spies.asp?error=6")
								end if
						end if
				end if
		end if
end if									
end if											
												
%>
<!---#include file="includes/end_check.asp"--->