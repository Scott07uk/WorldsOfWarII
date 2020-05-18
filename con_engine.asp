<!---#include file="includes/check.asp"--->
<%
'<!---SPELL CHECKED--->

Dim rscon			'varible to hold the collection
Dim intConNum		'varible to hold the number of the construction
Dim blnAllowed		'varible to decide if the construction is allowed
Dim strAttribCode	'varible to hold the code for the attribute in the database
Dim intTickTime		'varible to hold the time for the construction to take place
Dim lngFoodCost		'varible to hold the food cost
Dim lngWoodCost		'varible to hold the week cost
Dim lngIronCost		'varible to hold the iron cost
Dim lngMoneyCost	'varible to hold the money cost

'get the constructiuon number
intConNum = Request.QueryString("start")

'validate the input
if blnTickRunning = true then
		Response.Redirect("construction.asp?error=5")
	else
if isnumeric(intConNum) = false then
		'the input is invalid
		Response.Redirect("construction.asp")
	else
		if (intConNum < 1) OR (intconnum > 16) then
				'the input is invalid
				Response.Redirect("construction.asp")
			else
				'create the sql statement to query the database
				strSQL = "SELECT food, money, iron, wood, curcon, conticks, ResCE, ResVGM, ResSF, ResMPV, ResBGP, ResMWP, ResMPC, ResRP, ResFPC, ResEP, ConAA, ConJF, ConMF, ConSF, ConSPF, ConHMP, ConLF, ConHF, ConWF, ConBF, ResLB, ResFG, ResSIT, ConRF, ConMP, ResCS, ResE, ConDC, ConRS, ConSC, ConSRC FROM tblAccounts WHERE username = '" & strUsername & "';"
				
				'create the object to handle the collection
				set rscon = Server.CreateObject("ADODB.Recordset")
				
				rscon.LockType = 3
				
				rscon.CursorType = 2
				
				rscon.Open strsql, adocon
				
				'the number is valid check which construction they want
				Select case intConNum
					case 1
						strAttribCode = "JF"
						intTickTime = 12
						lngFoodCost = 800
						lngWoodCost = 250
						lngIronCost = 1500
						lngMoneyCost = 1250
						
						'prerequsit check
						if cint(rsCon("ResVGM")) = 3 then
							blnAllowed = True
						else
							blnAllowed = False
						end if
					case 2
						strAttribCode = "AA"
						intTickTime = 15
						lngFoodCost = 500
						lngWoodCost = 500
						lngIronCost = 750
						lngMoneyCost = 1250
						
						'prerequisits
						if cint(rsCon("ResSF")) = 3 then
								blnAllowed = True
							else
								blnAllowed = False
						end if
					case 3
						strAttribCode = "MF"
						intTickTime = 18
						lngFoodCost = 850
						lngWoodCost = 500
						lngIronCost = 1000
						lngMoneyCost = 1500
						
						'prerequisits
						if (cint(rsCon("ResMPV")) = 3) AND (cint(rscon("ResVGM")) = 3) then
								blnAllowed = True
							else
								blnAllowed = False
						end if
					case 4
						strAttribCode = "SF"
						intTickTime = 24
						lngFoodCost = 1000
						lngWoodCost = 500
						lngIronCost = 1500
						lngMoneyCost = 2000
						
						'prerequisits
						if (cint(rsCon("ResMPV")) = 3) AND (cint(rsCon("resVGM")) = 3) AND (cint(rsCon("ResBGP")) = 3) then
								blnAllowed = True
							else
								blnAllowed = False
						end if
					case 5
						strAttribCode = "SPF"
						intTickTime = 36
						lngFoodCost = 1000
						lngWoodCost = 750
						lngIronCost = 2500
						lngMoneyCost = 3500
						
						'prerequisits
						if (cint(rscon("ResMPC")) = 3) AND (cint(rscon("ResCE")) = 3) then
								blnAllowed = True
							else
								blnAllowed = False
						end if
					case 6
						strAttribCode = "RF"
						intTickTime = 20
						lngFoodCost = 900
						lngWoodCost = 1250
						lngIronCost = 1980
						lngMoneyCost = 1000
						
						'prerequists
						if cint(rscon("ResRP")) = 3 then
								blnAllowed = True
							else
								blnAllowed = False
						end if
					case 7
						strAttribCode = "HMP"
						intTickTime = 28
						lngFoodCost = 800
						lngWoodCost = 800
						lngIronCost = 1250
						lngMoneyCost = 1700
						
						'prerequisits
						if (cint(rscon("ResFPC")) = 3) AND (cint(rsCon("ResCE")) = 3) then 
								blnAllowed = True
							else
								blnAllowed = False
						end if
					case 8
						strAttribCode = "HF"
						intTickTime = 32
						lngFoodCost = 1250
						lngWoodCost = 760
						lngIronCost = 1300
						lngMoneyCost = 2800
						
						'prerequisits
						if (cint(rsCon("ResMPC")) = 3) AND (cint(rscon("ResCE")) = 3) AND (cint(rscon("ConBF")) = 3) then
								blnAllowed = True
							else
								blnAllowed = False
						end if
					case 9	
						strAttribCode = "WF"
						intTickTime = 32
						lngFoodCost = 1250
						lngWoodCost = 1200
						lngIronCost = 2100
						lngMoneyCost = 3500
						
						'prerequistis
						if (cint(rsCon("ResMPC")) = 3) AND (cint(rsCon("ResLB")) = 3) AND (cint(rsCon("ResCE")) = 3) then
								blnAllowed = True
							else
								blnAllowed = False
						end if
					case 10
						strAttribCode = "LF"
						intTickTime = 48
						lngFoodCost = 1500
						lngWoodCost = 1300
						lngIronCost = 3300
						lngMoneyCost = 4250
						
						'prerequisits
						if (cint(rsCon("ResMPC")) = 3) AND (cint(rsCon("ResLB")) = 3) AND (cint(rsCon("ResCE")) = 3) then
								blnAllowed = False
							else
								blnAllowed = True
						end if
					case 11
						strAttribCode = "MP"
						intTickTime = 72
						lngFoodCost = 3500
						lngWoodCost = 2500
						lngIronCost = 4500
						lngMoneyCost = 7800
						
						'prerequsits
						blnAllowed = True
					case 12
						strAttribCode = "DC"
						intTickTime = 24
						lngFoodCost = 1750
						lngWoodCost = 750
						lngIronCost = 1500
						lngMoneyCost = 2500
						
						'prerequisits
						blnAllowed = True
					case 13
						strAttribCode = "RS"
						intTickTime = 15
						lngFoodCost = 250
						lngWoodCost = 500
						lngIronCost = 1250
						lngMoneyCost = 1500
						
						'prerequisits
						
						if cint(rscon("resFG")) = 3 then
								blnAllowed = True
							else
								blnAllowed = False
						end if
					case 15
						strAttribCode = "SC"
						intTickTime = 24
						lngFoodCost = 1500
						lngWoodCost = 600
						lngIronCost = 2000
						lngMoneyCost = 2500
						
						'prerequsits
						if (cint(rsCon("ConAA")) = 3) AND (cint(rsCon("ResSIT")) = 3) AND (cint(rsCon("ConMP")) = 3) AND (cint(rscon("ResRP")) = 3) AND (cint(rscon("ResCS")) = 3) AND (cint(rscon("ResE")) = 3) then
								blnAllowed = True
							else
								blnAllowed = False
						end if
					case 16
						strAttribCode = "SRC"
						intTickTime = 48
						lngFoodCost = 2500
						lngWoodCost = 1500
						lngIronCost = 3500
						lngMoneyCost = 5000
						
						'prerequisits
						if cint(rscon("conSRC")) = 3 then
								blnAllowed = True
							else
								blnAllowed = False
						end if
				end select
				
				if clng(rscon("money")) < lngMoneyCost then
						Response.Redirect("construction.asp?error=1")
					else
						if clng(rscon("iron")) < lngironCost then
								Response.Redirect("construction.asp?error=2")
							else
								if clng(rscon("wood")) < lngWoodCost then
										Response.Redirect("construction.asp?error=3")
									else
										if clng(rscon("food")) < lngfoodcost then
												Response.Redirect("construction.asp?error=4")
											else
												if rscon("curcon") <> "" then
														Response.Redirect("construction.asp?error=5")
													else
														'start the construction
														
														'debit the resorces
														call rscon.Update("money", clng(rscon("money")) - lngmoneycost)
														call rscon.Update("iron", clng(rscon("iron")) - lngironcost)
														call rscon.Update("wood", clng(rscon("wood")) - lngwoodcost)
														call rscon.Update("food", clng(rscon("food")) - lngfoodcost)
														
														'save the construction details
														call rscon.Update("curcon", strAttribCode)
														call rscon.Update("con" & strAttribCode,2)
														call rscon.Update("conticks", intTickTime)
														
														Response.Redirect("construction.asp?error=0")
												end if
										end if
								end if
						end if
				end if
				
				adocon.close
				
				set adocon = nothing
		end if
end if
end if
%>

<!---#include file="includes/end_check.asp"--->