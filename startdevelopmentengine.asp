<%
option explicit
%>
<!---#include file="includes/check.asp"--->

<%
'SPELL CHECKED

dim rsIndustry
Dim lngOrderNo
Dim lngFarmOrder
Dim lngMillOrder
Dim lngMineOrder
Dim lngrefineryOrder
Dim lngFarmTotal
Dim lngMilltotal
Dim lngRefineryTotal
Dim lngMineTotal
Dim lngFoodCost
Dim lngWoodCost
Dim lngIronCost
Dim lngMoneyCost

if isnumeric(Request.Form("farms")) = false then
		lngFarmOrder = 0
	else
		lngFarmOrder = clng(Request.form("farms"))
end if
if isnumeric(Request.form("sawmills")) = false then
		lngMillOrder = 0
	else
		lngmillOrder = clng(Request.form("sawmills"))
end if
if isnumeric(Request.Form("refineries")) = false then
		lngRefineryOrder = 0
	else
		lngRefineryorder = clng(Request.form("refineries"))
end if
if isnumeric(Request.Form("mines")) = false then
		lngMineOrder = 0
	else
		lngMineOrder = clng(Request.Form("mines"))
end if

If lngfarmorder < 0 then
		lngFarmOrder = 0
end if
if lngmillorder < 0 then
		lngmillorder = 0
end if
if lngrefineryorder < 0 then
		lngrefineryorder = 0
end if
if lngmineorder < 0 then
		lngmineorder = 0
end if

lngFarmTotal = lngfarmorder
lngmilltotal = lngmillorder
lngRefineryTotal = lngrefineryorder
lngMineTotal = lngmineorder
if blnTickRunning = true then
		Response.Redirect("development.asp?error=13")
	else
IF Request.Form("farms") = 0 and Request.Form("sawmills") = 0 and Request.Form("refineries") = 0 and Request.Form("mines") = 0 then
		Response.Redirect("development.asp?error=1")
	else
		set rsIndustry = server.CreateObject("adodb.recordset")

		rsIndustry.LockType = 2
	
		rsIndustry.CursorType = 3
	
		strSQL = "SELECT numfarms, nummills, numrefineries, nummines FROM tblAccounts WHERE username='" & strUsername & "';"
	
		rsIndustry.Open StrSQL, adocon
	
		lngFarmTotal = lngFarmTotal + clng(rsIndustry("numfarms"))
		lngmillTotal = lngmilltotal + clng(rsindustry("nummills"))
		lngrefinerytotal = lngrefinerytotal + clng(rsindustry("numrefineries"))
		lngminetotal = lngminetotal + clng(rsindustry("nummines"))
		
		rsIndustry.Close 
		
		strsql = "SELECT farms, mines, refineries, sawmills FROM tblpendingbuildings WHERE username = '" & strusername & "';"
		
		rsIndustry.Open strsql, adocon
		
		do while not rsIndustry.EOF 
			lngfarmtotal = lngfarmtotal + clng(rsindustry("farms"))
			lngmillTotal = lngmilltotal + clng(rsindustry("sawmills"))
			lngrefinerytotal = lngrefinerytotal + clng(rsindustry("refineries"))
			lngminetotal = lngminetotal + clng(rsindustry("mines"))
			
			rsIndustry.MoveNext 
		loop
		
		if clng(lngfarmtotal+lngmilltotal+lngrefinerytotal+lngminetotal) > lngland then
				Response.Redirect("development.asp?error=2")
			else
				'calculate the cost
				
				lngFoodCost = clng(lngfarmorder * 250)
				lngWoodCost = clng(lngfarmorder * 100)
				lngIronCost = clng(lngfarmorder * 100)
				lngMoneyCost = clng(lngfarmorder * 100)


				lngFoodCost = lngFoodCost + clng(lngmillorder * 100)
				lngWoodCost = lngWoodCost + clng(lngmillorder * 0)
				lngIronCost = lngIronCost + clng(lngmillorder * 250)
				lngMoneyCost = lngMoneycost + clng(lngmillorder * 500)
				
				lngFoodCost = lngFoodCost + clng(lngrefineryorder * 100)
				lngWoodCost = lngWoodCost + clng(lngrefineryorder * 0)
				lngIronCost = lngIronCost + clng(lngrefineryorder * 500)
				lngMoneyCost = lngMoneycost + clng(lngrefineryorder * 750)
				
				lngFoodCost = lngFoodCost + clng(lngmineorder * 500)
				lngWoodCost = lngWoodCost + clng(lngmineorder * 50)
				lngIronCost = lngIronCost + clng(lngmineorder * 150)
				lngMoneyCost = lngMoneycost + clng(lngmineorder * 200)
				
				if lngFoodCost > lngFood then
						Response.Redirect("development.asp?error=3")
					else
						if lngWoodCost > lngWood then
								Response.Redirect("development.asp?error=4")
							else
								if lngironCost > lngiron then
										Response.Redirect("development.asp?error=5")
									else
										if lngmoneycost > lngmoney then
												Response.Redirect("development.asp?error=6")
											else
												rsIndustry.Close 
												
												strsql = "SELECT food, wood, iron, money FROM tblaccounts WHERE username = '" & strusername & "';"
												
												rsIndustry.Open strsql, adocon
												
												with rsindustry 
													.Fields("food") = clng(rsindustry("food")) - lngfoodCost
													.Fields("wood") = clng(rsindustry("wood")) - lngwoodcost
													.Fields("iron") = clng(rsindustry("iron")) - lngironcost
													.Fields("money") = clng(rsindustry("money")) - lngmoneycost
													
													.Update
												end with
												
												rsIndustry.Close 
												
												strsql = "SELECT orderid FROM tblpendingbuildings ORDER BY orderid DESC;"
												
												rsIndustry.Open strsql, adocon
												
												if rsIndustry.EOF then
														lngOrderNo = 1 
													else
														lngOrderNo = clng(rsindustry("orderid")) + 1
												end if
												
												if lngmineorder > 0 then
														rsIndustry.Close 
														
														strsql = "SELECT orderid, ticksleft, mines, username FROM tblpendingbuildings WHERE ticksleft = 6 AND username = '" & strusername & "';"
														
														rsIndustry.Open strsql, adocon
														
														if rsIndustry.eof then
																with rsindustry 
																	.addnew
																	.fields("orderid") = lngOrderNo
																	.fields("ticksleft") = 6
																	.fields("mines") = lngmineorder
																	.fields("username") = strusername
																	
																	.update
																end with
																lngOrderNo = lngorderno + 1
															else
																with rsindustry
																	.Fields("mines") = clng(rsindustry("mines")) + lngmineorder
																	.update
																end with
														end if
												end if
												
												if lngfarmorder > 0 then
														rsIndustry.Close 
														
														strsql = "SELECT orderid, ticksleft, farms, username FROM tblpendingbuildings WHERE ticksleft = 8 AND username = '" & strusername & "';"
														
														rsIndustry.Open strsql, adocon
														
														if rsIndustry.eof then
																with rsindustry 
																	.addnew
																	.fields("orderid") = lngOrderNo
																	.fields("ticksleft") = 8
																	.fields("farms") = lngfarmorder
																	.fields("username") = strusername
																	
																	.update
																end with
																lngOrderNo = lngorderno + 1
															else
																with rsindustry
																	.Fields("farms") = clng(rsindustry("farms")) + lngfarmorder
																	.update
																end with
														end if
												end if
												
												if (lngmillorder > 0) OR (lngrefineryorder > 0) then
														rsIndustry.Close 
														
														strsql = "SELECT orderid, ticksleft, sawmills, refineries, username FROM tblpendingbuildings WHERE ticksleft = 10 AND username = '" & strusername & "';"
														
														rsIndustry.Open strsql, adocon
														
														if rsIndustry.eof then
																with rsindustry 
																	.addnew
																	.fields("orderid") = lngOrderNo
																	.fields("ticksleft") = 10
																	.fields("sawmills") = lngmillorder
																	.fields("refineries") = lngrefineryorder
																	.fields("username") = strusername
																	
																	.update
																end with
																lngOrderNo = lngorderno + 1
															else
																with rsindustry
																	.Fields("sawmills") = clng(rsindustry("sawmills")) + lngmillorder
																	.Fields("refineries") = clng(rsindustry("refineries")) + lngrefineryorder
																	.update
																end with
														end if
												end if
												
												Response.Redirect("development.asp?error=7")
										end if
								end if
						end if
				end if
		end if
		
		rsIndustry.close 
		
		set rsindustry = nothing

end if
end if
%>

<!---#include file="includes/end_check.asp"--->
