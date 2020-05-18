<%
'SPELL CHECKED

Option Explicit
%>
<!---#include file="includes\check.asp"--->
<%

'Declair varibles
Dim RsOrder			'collection used to handle the order
Dim lngInfantry		'the number of infantry to be ordered
Dim lngComandoes	'the number of comandoes to be orderd
Dim lngJeeps		'the number of jeeps to be ordered
Dim lngMatilda1		'the number of matilda mk1 to be ordered
Dim lngMatilda2		'the number of matilida mk2 to be ordered
Dim lngSherman		'the number of shermans to be ordered
Dim lngSpitfire1	'the number of spitfire 1's to be ordered
Dim lngSpitfire2	'"
Dim lngSpitfire5	'"
Dim lngHurricane	'the number of hurricanes to be ordered
Dim lngHalifax		'the number of halifax on the order
dim lngWellington	'the number of wellingstos to be ordered
Dim lngLancaster	'the number of lancs to be ordered
Dim lngLandingCraft	'the number of landing craft to be ordered
Dim lngLandingship	'the number of tlaning ships to be ordered
Dim lngFrigates		'the number of frigates to be ordered
Dim lngBattleships	'the number of battleships to be ordered
Dim lngCarriers		'the number of aircraft carriers to be ordered
Dim lngFoodCost		'the food cost of the order
Dim lngWoodCost		'the wood cost of the order
Dim lngIronCost		'the iron cost of the order
Dim lngMoneyCost	'the money cost of the order
Dim lngOrderNo

'get the input from the user

lngInfantry = Request.Form("infantry")
lngComandoes = Request.Form("comandoes")
lngJeeps = Request.Form("Jeeps")
lngMatilda1 = Request.Form("matilda1")
lngmatilda2 = Request.Form("matilda2")
lngsherman = Request.Form("sherman")
lngspitfire1 = Request.Form("spitfire1")
lngspitfire2 = Request.Form("spitfire2")
lngspitfire5 = Request.Form("spitfire5")
lngHurricane = Request.Form("hurricane")
lnghalifax = Request.Form("halifax")
lngwellington = Request.Form("wellington")
lnglancaster = Request.form("lancaster")
lnglandingcraft = Request.Form("landingcraft")
lnglandingship = Request.Form("landingship")
lngfrigates = Request.Form("frigates")
lngbattleships = Request.Form("battleship")
lngcarriers = Request.Form("carriers")


'validate the user input

if (isnumeric(lnginfantry) = false) OR (lnginfantry="") then
		lnginfantry = 0
end if
if (isnumeric(lngcomandoes) = false) OR (lngcomandoes="") then
		lngcomandoes = 0
end if
if (isnumeric(lngjeeps) = false) OR (lngjeeps="") then
		lngjeeps = 0
end if
if (isnumeric(lngmatilda1) = false) OR (lngmatilda1="") then
		lngmatilda1 = 0
end if
if (isnumeric(lngmatilda2) = false) OR (lngmatilda2="") then
		lngmatilda2 = 0
end if
if (isnumeric(lngsherman) = false) OR (lngsherman="") then
		lngsherman = 0
end if
if (isnumeric(lngspitfire1) = false) or (lngspitfire1 = "") then
		lngspitfire1 = 0
end if
if (isnumeric(lngspitfire2) = false) OR (lngspitfire2 = "") then
		lngspitfire2 = 0
end if
if (isnumeric(lngspitfire5) = false) OR (lngspitfire5="") then
		lngspitfire5 = 0
end if
if (isnumeric(lnghurricane) = false) OR (lnghurricane="") then
		lnghurricane = 0
end if
if (isnumeric(lnghalifax) = false) OR (lnghalifax="") then
		lnghalifax = 0
end if
if (isnumeric(lngwellington) = false) OR (lngwellington="") then
		lngwellington = 0
end if
if (isnumeric(lnglancaster) = false) OR (lnglancaster="") then
		lnglancaster = 0
end if
if (isnumeric(lnglandingcraft) = false) OR (lnglandingcraft="") then
		lnglandingcraft = 0
end if
if (isnumeric(lnglandingship) = false) OR (lnglandingship="") then
		lnglandingship = 0
end if
if (isnumeric(lngfrigates) = false) OR (lngfrigates="") then
		lngfrigates = 0
end if
if (isnumeric(lngbattleships) = false) OR (lngbattleships="") then
		lngbattleships = 0
end if
if (isnumeric(lngcarriers) = false) OR (lngcarriers="") then
		lngcarriers = 0
end if

if lnginfantry < 0 then
		lnginfantry = 0
end if
if lngcomandoes < 0 then
		lngcomandoes = 0
end if
if lngmatilda1 < 0 then
		lngmatlida1 = 0
end if
if lngmatilda2 < 0 then
		lngmatilda2 = 0
end if
if lngjeeps < 0 then
		lngjeeps = 0
end if
if lngsherman < 0 then
		lngsherman = 0
end if
if lngspitfire1 < 0 then
		lngspitfire1 = 0
end if
if lngspitfire2 < 0 then
		lngspitfire2 = 0
end if
if lngspitfire5 < 0 then
		lngspitfire5 = 0
end if
if lnghurricane < 0 then
		lnghurricane = 0
end if
if lngHalifax < 0 then
		lnghalifax = 0
end if
if lngwellington < 0 then
		lngwellington = 0
end if
if lnglancaster < 0 then
		lnglancaster = 0
end if
if lnglandingcraft < 0 then
		lnglandingcraft = 0
end if
if lnglandingship < 0 then
		lnglandingship = 0
end if
if lngfrigates < 0 then
		lngfrigates = 0
end if
if lngbattleships < 0 then
		lngbattleships = 0
end if
if lngcarriers < 0 then
		lngcarriers = 0
end if

if blnTickRunning = true then
		Response.Redirect("units.asp?error=4")
	else
if (lnginfantry+lngcomandoes+lngjeeps+lngmatilda1+lngmatilda2+lngsherman+lngspitfire1+lngspitfire2+lngspitfire5+lnghurricane+lnghalifax+lngwellington+lnglancaster+lnglandingcraft+lnglandingship+lngfrigates+lngbattleships+lngcarriers) = 0 then
		Response.Redirect("units.asp?error=1")
	else
		'there is stuff to order
		
		'calculate the cost of the order
		lngfoodCost = 0
		lngironcost = 0
		lngwoodcost = 0
		lngmoneycost = 0
		
		'add the infantry cost
		lngfoodcost = lnginfantry * 50
		lngironcost = lnginfantry * 5
		lngmoneycost = lnginfantry * 100
		
		'add the comandoes cost
		lngfoodcost = lngfoodcost + (lngcomandoes * 75)
		lngironcost = lngironcost + (lngcomandoes * 10)
		lngmoneycost = lngmoneycost + (lngcomandoes * 150)
		
		'add the jeeps cost
		lngfoodcost = lngfoodcost + (lngjeeps * 100)
		lngironcost = lngironcost + (lngjeeps * 75)
		lngmoneycost = lngmoneycost + (lngjeeps * 300)
		
		'add the cost for matilda mk1
		lngfoodcost = lngfoodcost + (lngmatilda1 * 150)
		lngironcost = lngironcost + (lngmatilda1 * 100)
		lngmoneycost = lngmoneycost + (lngmatilda1 * 500)
		
		'add the cost for matilda mk2
		lngfoodcost = lngfoodcost + (lngmatilda2 * 150)
		lngironcost = lngironcost + (lngmatilda2 * 120)
		lngmoneycost = lngmoneycost + (lngmatilda2 * 600)
		
		'add the cost for sherman
		lngfoodcost = lngfoodcost + (lngsherman * 150)
		lngironcost = lngironcost + (lngsherman * 150)
		lngmoneycost = lngmoneycost + (lngsherman * 700)
		
		'add the cost for spitfire mk1
		lngfoodcost = lngfoodcost + (lngspitfire1 * 200)
		lngwoodcost = lngwoodcost + (lngspitfire1 * 10)
		lngironcost = lngironcost + (lngspitfire1 * 100)
		lngmoneycost = lngmoneycost + (lngspitfire1 * 700)
		
		'add the cost for spitfire mk2
		lngfoodcost = lngfoodcost + (lngspitfire2 * 200)
		lngwoodcost = lngwoodcost + (lngspitfire2 * 10)
		lngironcost = lngironcost + (lngspitfire2 * 125)
		lngmoneycost = lngmoneycost + (lngspitfire2 * 750)
		
		'add the cost for spitfire mk5
		lngfoodcost = lngfoodcost + (lngspitfire5 * 250)
		lngwoodcost = lngwoodcost + (lngspitfire5 * 10)
		lngironcost = lngironcost + (lngspitfire5 * 150)
		lngmoneycost = lngmoneycost + (lngspitfire5 * 800)
		
		'add the cost for the hurricane
		lngfoodcost = lngfoodcost + (lnghurricane * 200)
		lngwoodcost = lngwoodcost + (lnghurricane * 50)
		lngironcost = lngironcost + (lnghurricane * 75)
		lngmoneycost = lngmoneycost + (lnghurricane * 600)
		
		'add the cost for the halifax bomber
		lngfoodcost = lngfoodcost + (lnghalifax * 300)
		lngironcost = lngironcost + (lnghalifax * 200)
		lngmoneycost = lngmoneycost + (lnghalifax * 900)
		
		'add the cost for the wellington
		lngfoodcost = lngfoodcost + (lngwellington * 300)
		lngironcost = lngironcost + (lngwellington * 200)
		lngmoneycost = lngmoneycost + (lngwellington * 1000)
		
		'add the cost for the lancaster
		lngfoodcost = lngfoodcost + (lnglancaster * 350)
		lngironcost = lngironcost + (lnglancaster * 300)
		lngmoneycost = lngmoneycost + (lnglancaster * 1100)
		
		'add the cost for the landing craft
		lngfoodcost = lngfoodcost + (lnglandingcraft * 75)
		lngironcost = lngironcost + (lnglandingcraft * 90)
		lngmoneycost = lngmoneycost + (lnglandingcraft * 150)
		
		'add the cost for the landing ships
		lngfoodcost = lngfoodcost + (lnglandingship * 150)
		lngwoodcost = lngwoodcost + (lnglandingship * 10)
		lngironcost = lngironcost + (lnglandingship * 500)
		lngmoneycost = lngmoneycost + (lnglandingship * 300)
		
		'add the cost for the frigates
		lngfoodcost = lngfoodcost + (lngfrigates * 250)
		lngwoodcost = lngwoodcost + (lngfrigates * 20)
		lngironcost = lngironcost + (lngfrigates * 1000)
		lngmoneycost = lngmoneycost + (lngfrigates * 1100)
		
		'add the cost for the battleships
		lngfoodcost = lngfoodcost + (lngbattleships * 400)
		lngwoodcost = lngwoodcost + (lngbattleships * 30)
		lngironcost = lngironcost + (lngbattleships * 2000)
		lngmoneycost = lngmoneycost + (lngbattleships * 1500)
		
		'add the cost for the carriers
		lngfoodcost = lngfoodcost + (lngcarriers * 500)
		lngwoodcost = lngwoodcost + (lngcarriers * 40)
		lngironcost = lngironcost + (lngcarriers * 3000)
		lngmoneycost = lngmoneycost + (lngcarriers * 2000)
		
		if (lngfoodcost > lngfood) OR (lngwoodcost > lngwood) OR (lngironcost > lngiron) OR (lngmoneycost > lngmoney) then
				Response.Redirect("units.asp?error=2")
			else
				
				'create the objects to perform the work
				
				set rsOrder = server.CreateObject("ADODB.Recordset")
				
				RsOrder.LockType = 3
				RsOrder.CursorType = 2
				
				'generate the sql to take the resorces
				strsql = "SELECT food, wood, money, iron FROM tblaccounts WHERE username = '" & strusername & "';"
				
				RsOrder.Open strsql, adocon
				
				RsOrder.Fields("food") = clng(rsorder("food")) - lngfoodcost
				RsOrder.Fields("wood") = clng(rsorder("wood")) - lngwoodcost
				RsOrder.Fields("money") = clng(rsorder("money")) - lngmoneycost
				RsOrder.Fields("iron") = clng(rsorder("iron")) - lngironcost
				
				RsOrder.Update 
				
				RsOrder.Close 
				
				strsql = "SELECT orderno FROM tblpendingunits ORDER BY orderno DESC;"
				
				RsOrder.Open strsql, adocon
				
				if RsOrder.EOF then
						lngorderno = 1
					else
						lngorderNo = clng(rsorder("orderno"))
				end if
				
				RsOrder.Close 
				
				lngorderno = lngorderno + 1
				
				if lnginfantry > 0 then
						strsql = "SELECT username, timeleft, infantry, orderno FROM tblpendingunits WHERE timeleft = 3 AND username = '" & strusername & "';"
						RsOrder.Open strsql, adocon
						
						if RsOrder.EOF then 
								with rsorder
									.addnew
									.fields("username") = strusername
									.fields("timeleft") = 3
									.fields("infantry") = lnginfantry
									.fields("orderno") = lngorderno
									.update
								end with
								lngorderno = lngorderno + 1
							else
								call RsOrder.Update("infantry", clng(rsorder("infantry")) + lnginfantry)
						end if
						
						RsOrder.Close 
				end if
				
				if lngcomandoes > 0 then
						strsql = "SELECT username, timeleft, comandoes, orderno FROM tblpendingunits WHERE timeleft = 4 AND username = '" & strusername & "';"
						RsOrder.Open strsql, adocon
						
						if RsOrder.EOF then
								with RsOrder
									.addnew
									.fields("username") = strusername
									.fields("timeleft") = 4
									.fields("comandoes") = lngcomandoes
									.fields("orderno") = lngorderno
									.update
								end with
								lngorderno = lngorderno + 1
							else
								call RsOrder.Update("comandoes", clng(rsorder("comandoes")) + lngcomandoes)
						end if
						
						RsOrder.Close 
				end if
				
				if lnglandingcraft > 0 then
						strsql = "SELECT username, timeleft, landingcraft, orderno FROM tblpendingunits WHERE timeleft = 5 AND username = '" & strusername & "';"
						RsOrder.Open strsql, adocon
						
						if RsOrder.EOF then
								with RsOrder
									.addnew
									.fields("username") = strusername
									.fields("timeleft") = 5
									.fields("landingcraft") = lnglandingcraft
									.fields("orderno") = lngorderno
									.update
								end with
								lngorderno = lngorderno + 1
							else
								call RsOrder.Update("landingcraft", clng(rsorder("landingcraft")) + lnglandingcraft)
						end if
						
						RsOrder.Close 
				end if
				
				if lngjeeps > 0 then
						strsql = "SELECT username, timeleft, jeeps, orderno FROM tblpendingunits WHERE timeleft = 6 AND username = '" & strusername & "';"
						RsOrder.Open strsql, adocon
						
						if RsOrder.EOF then
								with RsOrder
									.addnew
									.fields("username") = strusername
									.fields("timeleft") = 6
									.fields("jeeps") = lngjeeps
									.fields("orderno") = lngorderno
									.update
								end with
								lngorderno = lngorderno + 1
							else
								call RsOrder.Update("jeeps", clng(rsorder("jeeps")) + lngjeeps)
						end if
						
						RsOrder.Close 
				end if
				
				if lnghurricane > 0 then
						strsql = "SELECT username, timeleft, hurricane, orderno FROM tblpendingunits WHERE timeleft = 7 AND username = '" & strusername & "';"
						RsOrder.Open strsql, adocon
						
						if RsOrder.EOF then
								with RsOrder
									.addnew
									.fields("username") = strusername
									.fields("timeleft") = 7
									.fields("hurricane") = lnghurricane
									.fields("orderno") = lngorderno
									.update
								end with
								lngorderno = lngorderno + 1
							else
								call RsOrder.Update("hurricane", clng(rsorder("hurricane")) + lnghurricane)
						end if
						
						RsOrder.Close 
				end if
				
				if (lngmatilda1 > 0) OR (lngmatilda2 > 0) OR (lngspitfire1 > 0) then
						strsql = "SELECT username, timeleft, matilda1, matilda2, spitfire1, orderno FROM tblpendingunits WHERE timeleft = 8 AND username = '" & strusername & "';"
						RsOrder.Open strsql, adocon
						
						if RsOrder.EOF then
								with RsOrder
									.addnew
									.fields("username") = strusername
									.fields("timeleft") = 8
									.fields("matilda1") = lngmatilda1
									.fields("matilda2") = lngmatilda2
									.fields("spitfire1") = lngspitfire1
									.fields("orderno") = lngorderno
									.update
								end with
								lngorderno = lngorderno + 1
							else
								with rsorder
									.Fields("matilda1") = clng(rsorder("matilda1")) + lngmatilda1
									.Fields("matilda2") = clng(rsorder("matilda2")) + lngmatilda2
									.Fields("spitfire1") = clng(rsorder("spitfire1")) + lngspitfire1
									.Update
								end with
						end if
						
						RsOrder.Close 
				end if
				
				if (lngsherman > 0) OR  (lngspitfire2 > 0) OR (lnghalifax > 0) OR (lnglandingship > 0) then
						strsql = "SELECT username, timeleft, sherman, spitfire2, halifax, landingship, orderno FROM tblpendingunits WHERE timeleft = 10 AND username = '" & strusername & "';"
						RsOrder.Open strsql, adocon
						
						if RsOrder.EOF then
								with RsOrder
									.addnew
									.fields("username") = strusername
									.fields("timeleft") = 10
									.fields("sherman") = lngsherman
									.fields("spitfire2") = lngspitfire2
									.fields("halifax") = lnghalifax
									.fields("landingship") = lnglandingship
									.fields("orderno") = lngorderno
									.update
								end with
								lngorderno = lngorderno + 1
							else
								with rsorder
									.Fields("sherman") = clng(rsorder("sherman")) + lngsherman
									.Fields("spitfire2") = clng(rsorder("spitfire2")) + lngspitfire2
									.Fields("halifax") = clng(rsorder("halifax")) + lnghalifax
									.Fields("landingship") = clng(rsorder("landingship")) + lnglandingship
									
									.Update
								end with
						end if
						
						RsOrder.Close 
				end if
				
				
				if (lngspitfire5 > 0) OR (lngwellington > 0) then
						strsql = "SELECT username, timeleft, spitfire5, wellington, orderno FROM tblpendingunits WHERE timeleft = 12 AND username = '" & strusername & "';"
						RsOrder.Open strsql, adocon
						
						if RsOrder.EOF then
								with RsOrder
									.addnew
									.fields("username") = strusername
									.fields("timeleft") = 12
									.fields("spitfire5") = lngspitfire5
									.fields("wellington") = lngwellington
									.fields("orderno") = lngorderno
									.update
								end with
								lngorderno = lngorderno + 1
							else
								with RsOrder
									.Fields("spitfire5") = clng(rsorder("spitfire5")) + lngspitfire5
									.Fields("wellington") = clng(rsorder("wellington")) + lngwellington
									
									.Update 
								end with
						end if
						
						RsOrder.Close 
				end if
				
				if lnglancaster > 0 then
						strsql = "SELECT username, timeleft, lancaster, orderno FROM tblpendingunits WHERE timeleft = 15 AND username = '" & strusername & "';"
						RsOrder.Open strsql, adocon
						
						if RsOrder.EOF then
								with RsOrder
									.addnew
									.fields("username") = strusername
									.fields("timeleft") = 15
									.fields("lancaster") = lnglancaster
									.fields("orderno") = lngorderno
									.update
								end with
								lngorderno = lngorderno + 1
							else
								with RsOrder
									.Fields("lancaster") = clng(rsorder("lancaster")) + lnglancaster
									
									.Update 
								end with
						end if
						
						RsOrder.Close 
				end if
				
				
				if lngfrigates > 0 then
						strsql = "SELECT username, timeleft, frigates, orderno FROM tblpendingunits WHERE timeleft = 16 AND username = '" & strusername & "';"
						RsOrder.Open strsql, adocon
						
						if RsOrder.EOF then
								with RsOrder
									.addnew
									.fields("username") = strusername
									.fields("timeleft") = 16
									.fields("frigates") = lngfrigates
									.fields("orderno") = lngorderno
									.update
								end with
								lngorderno = lngorderno + 1
							else
								with RsOrder
									.Fields("frigates") = clng(rsorder("frigates")) + lngfrigates
									
									.Update 
								end with
						end if
						
						RsOrder.Close 
				end if
				
				if lngbattleships > 0 then
						strsql = "SELECT username, timeleft, battleship, orderno FROM tblpendingunits WHERE timeleft = 20 AND username = '" & strusername & "';"
						RsOrder.Open strsql, adocon
						
						if RsOrder.EOF then
								with RsOrder
									.addnew
									.fields("username") = strusername
									.fields("timeleft") = 20
									.fields("battleship") = lngbattleships
									.fields("orderno") = lngorderno
									.update
								end with
								lngorderno = lngorderno + 1
							else
								with RsOrder
									.Fields("battleship") = clng(rsorder("battleship")) + lngbattleships
									
									.Update 
								end with
						end if
						
						RsOrder.Close 
				end if
				
				if lngcarriers > 0 then
						strsql = "SELECT username, timeleft, carrier, orderno FROM tblpendingunits WHERE timeleft = 24 AND username = '" & strusername & "';"
						RsOrder.Open strsql, adocon
						
						if RsOrder.EOF then
								with RsOrder
									.addnew
									.fields("username") = strusername
									.fields("timeleft") = 24
									.fields("carrier") = lngcarriers
									.fields("orderno") = lngorderno
									.update
								end with
								lngorderno = lngorderno + 1
							else
								with RsOrder
									.Fields("carrier") = clng(rsorder("carrier")) + lngcarriers
									
									.Update 
								end with
						end if
						
						RsOrder.Close 
				end if
				
				set rsorder = nothing
				
				Response.Redirect("units.asp?error=3")
		end if
end if
end if

%>
<!---#include file="includes\end_check.asp"--->