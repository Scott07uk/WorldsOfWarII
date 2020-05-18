<!---#include file="includes/check.asp"--->

<!---SPELL CHECKED--->

<%
'Define varibles
Dim rsExplore
Dim lngOrder
Dim lngOrderNo
Dim lngMoneyCost
Dim lngFoodCost
Dim lngWoodCost
Dim lngIronCost

'get the user input
lngOrder = Request.Form("explorers")

'validate the user input
lngOrder = trim(lngOrder)
if blnTickRunning = True then
		Response.Redirect("exploer.asp?error=12")
	else
if isnumeric(lngOrder) = False then
		Response.Redirect("explore.asp?error=1")
	else
		if lngOrder < 1 then
				Response.Redirect("explore.asp?error=2")
			else
				'generate the costs
				lngmoneyCost = lngorder * 500
				lngFoodCost = lngOrder * 250
				lngWoodCost = lngOrder * 10
				lngIronCost = lngOrder * 50
				
				'make sure the user has enough resorces
				if clng(lngmoneyCost) > lngMoney then
						Response.Redirect("explore.asp?error=3")
					else
						if lngFoodCost > lngFood then
								Response.Redirect("explore.asp?error=4")
							else
								if lngWoodCost > lngWood then
										Response.Redirect("explore.asp?error=5")
									else
										if lngIronCost > lngIron then
												Response.Redirect("explore.asp?error=6")
											else
											
												strSQL = "SELECT explorers FROM tblpendingunits WHERE timeleft = 6 AND username = '" & strusername & "';"
				
												'create the objects
												set rsExplore = server.CreateObject("ADODB.Recordset")
				
												'set the record attributes
												rsExplore.LockType = 3
												rsExplore.CursorType = 2
				
												rsExplore.Open strsql, adocon
				
												if rsExplore.EOF then
														'create a new record
														rsExplore.Close 
						
														strSQL = "SELECT * FROM tblpendingunits ORDER by orderno DESC;"
						
														rsExplore.Open strsql, adocon
						
														if rsExplore.EOF then
																lngOrderNo = 1
															else
																lngOrderNo = clng(rsexplore("orderno"))
																lngOrderNo = lngOrderNo + 1
														end if
						
														with rsExplore
															.AddNew
															.Fields("orderno") = lngOrderNo
															.Fields("username") = strusername
															.Fields("timeleft") = 6
															.Fields("explorers") = lngOrder
															.Update 
														end with
													else
														call rsExplore.Update("explorers", clng(rsExplore("explorers")) +lngOrder)
												end if
				
												rsExplore.Close
				
												strsql = "SELECT money, food, wood, iron FROM tblaccounts WHERE username = '" & strusername & "';"
				
												rsExplore.Open strsql, adocon
				
												call rsExplore.Update("money", clng(lngmoney-lngmoneycost))
												call rsExplore.Update("wood", clng(lngwood-lngwoodcost))
												call rsExplore.Update("food", clng(lngfood-lngfoodcost))
												call rsExplore.Update("iron", clng(lngiron-lngironcost))
				
												rsExplore.Close 
				
												set rsexplore = nothing
				
												Response.Redirect("explore.asp?error=7")
										end if
								end if
						end if
				end if
		end if
end if
end if								

%>
<!---#include file="includes/end_check.asp"--->