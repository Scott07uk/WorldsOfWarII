<%
'SPELL CHECKED

Option Explicit
%>
<!---#include file="includes\check.asp"--->
<%

'Declare variables
Dim RsOrder			'collection used to handle the ordered.
Dim lngPillboxes	'The number of Pillboxes to be ordered.
Dim lngAAGuns		'The number of AA Guns to be ordered.
Dim lngBalloons		'The number of Barrage Balloons to be ordered.
Dim lngMines		'The number of Mines to be ordered.
Dim lngTurrets		'The number of Turrets to be ordered.
Dim lngFoodCost		'The food cost of the order.
Dim lngWoodCost		'The wood cost of the order.
Dim lngIronCost		'The iron cost of the order.
Dim lngMoneyCost	'The money cost of the order.
Dim lngOrderNo

'Requesting values entered.

lngPillboxes = Request.Form("pillboxes")
lngAAGuns = Request.Form("aaguns")
lngBalloons = Request.Form("balloons")
lngMines = Request.Form("mines")
lngTurrets = Request.Form("turrets")

'Validation.

if (isnumeric(lngPillboxes) = false) OR (lngPillboxes="") then
		lngPillboxes = 0
end if
if (isnumeric(lngAAGuns) = false) OR (lngAAGuns="") then
		lngAAGuns = 0
end if
if (isnumeric(lngBalloons) = false) OR (lngBalloons="") then
		lngBalloons = 0
end if
if (isnumeric(lngMines) = false) OR (lngMines="") then
		lngMines = 0
end if
if (isnumeric(lngTurrets) = false) OR (lngTurrets="") then
		lngTurrets = 0
end if

if lngPillboxes < 0 then
		lngPillboxes = 0
end if
if lngAAGuns < 0 then
		lngAAGuns = 0
end if
if lngBalloons < 0 then
		lngBalloons = 0
end if
if lngMines < 0 then
		lngMines = 0
end if
if lngTurrets < 0 then
		lngTurrets = 0
end if

if blnTickRunning = true then
		Response.Redirect("defence.asp?error=1")
	else
if (lngPillboxes+lngAAGuns+lngBalloons+lngMines+lngTurrets) = 0 then
		Response.Redirect("defence.asp?error=1")
	else
		'Ordering.
		
		'Calculate the cost of the order.
		lngFoodCost = 0
		lngIronCost = 0
		lngWoodCost = 0
		lngMoneyCost = 0
		
		'Add the Pillbox cost.
		lngFoodCost = lngPillboxes * 100
		lngIronCost = lngPillboxes * 200
		lngMoneyCost = lngPillboxes * 500
		
		'Add the AA Gun cost.
		lngFoodCost = lngFoodCost + (lngAAGuns * 100)
		lngIronCost = lngIronCost + (lngAAGuns * 50)
		lngMoneyCost = lngMoneyCost + (lngAAGuns * 750)
		
		'Add the Barrage Balloon cost.
		lngFoodCost = lngFoodCost + (lngBalloons * 100)
		lngIronCost = lngIronCost + (lngBalloons * 50)
		lngMoneyCost = lngMoneyCost + (lngBalloons * 100)
		
		'Add the Water Mine cost.
		lngFoodCost = lngFoodCost + (lngMines * 50)
		lngIronCost = lngIronCost + (lngMines * 50)
		lngMoneyCost = lngMoneyCost + (lngMines * 100)
		
		'Add the Turret cost.
		lngFoodCost = lngFoodCost + (lngTurrets * 100)
		lngIronCost = lngIronCost + (lngTurrets * 50)
		lngMoneyCost = lngMoneyCost + (lngTurrets * 250)
		
		if (lngFoodCost > lngFood) OR (lngWoodCost > lngWood) OR (lngIronCost > lngIron) OR (lngMoneyCost > lngMoney) then
				Response.Redirect("defence.asp?error=2")
			else
				
				'Objects.
				
				set rsorder = server.CreateObject("ADODB.Recordset")
				
				rsorder.LockType = 3
				rsorder.CursorType = 2
				
				'SQL query.
				
				StrSQL = "SELECT Food, Wood, Iron, Money FROM tblAccounts WHERE username = '" & strusername & "';"
				
				rsorder.Open StrSQL, Adocon
				
				rsorder.Fields("Food") = clng(rsorder("Food")) - lngFoodCost
				rsorder.Fields("Wood") = clng(rsorder("Wood")) - lngWoodCost
				rsorder.Fields("Iron") = clng(rsorder("Iron")) - lngIronCost
				rsorder.Fields("Money") = clng(rsorder("Money")) - lngMoneyCost
				
				rsorder.Update 
				rsorder.Close 
				
				StrSQL = "SELECT OrderNo FROM tblPendingUnits ORDER BY OrderNo DESC;"
				
				rsorder.Open StrSQL, Adocon
				
				If rsorder.EOF Then
						lngOrderNo = 1
					Else
						lngOrderNo = clng(rsorder("OrderNo"))
				End If
				
				rsorder.Close
				
				lngOrderNo = lngOrderNo + 1
				
				If (lngPillboxes > 0) Or (lngBalloons > 0) Or (lngMines > 0) Then
						StrSQL = "SELECT Username, TimeLeft, Pilboxes, Balloons, Mines, OrderNo FROM tblPendingUnits WHERE TimeLeft = 3 AND Username = '" & strusername & "';"
						rsorder.Open StrSQL, Adocon
						
						if rsorder.EOF then 
								with rsorder
									.addnew
									.fields("Username") = strUsername
									.fields("TimeLeft") = 3
									.fields("Pilboxes") = lngPillboxes
									.fields("Balloons") = lngBalloons
									.fields("Mines") = lngMines
									.fields("OrderNo") = lngOrderNo
									.update
								end with
								lngOrderNo = lngOrderNo + 1
							else
								call rsorder.Update("Pilboxes", clng(rsorder("Pilboxes")) + lngPillboxes)
								call rsorder.Update("Balloons", clng(rsorder("Balloons")) + lngBalloons)
								call rsorder.Update("Mines", clng(rsorder("Mines")) + lngMines)
						end if
						
						rsorder.Close
				end if
				
				If (lngAAGuns > 0) Or (lngTurrets > 0) Then
						StrSQL = "SELECT Username, TimeLeft, AAGuns, Turrets, OrderNo FROM tblPendingUnits WHERE TimeLeft = 4 AND username = '" & strusername & "';"
						rsorder.Open StrSQL, Adocon
						
						if rsorder.EOF then
								with rsorder
									.addnew
									.fields("Username") = strUsername
									.fields("TimeLeft") = 4
									.fields("AAGuns") = lngAAGuns
									.fields("Turrets") = lngTurrets
									.fields("OrderNo") = lngOrderNo
									.update
								end with
								lngOrderNo = lngOrderNo + 1
							else
								call rsorder.Update("AAGuns", clng(rsorder("AAGuns")) + lngAAGuns)
								call rsorder.Update("Turrets", clng(rsorder("Turrets")) + lngTurrets)
						end if
						
						rsorder.Close 
				end if
				
				set rsorder = nothing
				
				Response.Redirect("defence.asp?error=3")
		end if
end if
end if
%>
<!---#include file="includes\end_check.asp"--->