<!---#include file="includes/check.asp"--->
<%
'SPELL CHECKED

dim check				'checks the research is complete before doing another work
dim intResID			'its an id OH YEAH
dim food				'yummy
dim wood				'not very tasty unless you are a beaver
dim iron				'a bit hard to eat and not very tasty
dim money				'can be used to buy food, not iron or wood as they are not very edible
dim ticks				'normally found on dogs
dim Research			'its stuff
dim rsResearch
dim strResearch			'its stuff

intResID = request.querystring("ID")

'check that the ID is within acceptable range
if intResID < 1 or intResID > 25 then
	Response.redirect("research.asp")
end if
if blnTickRunning = true then
		Response.Redirect("research.asp?error=1")
	else
select case intResID
	case 1
		strresearch = "resCE"
	case 2
		strresearch = "resVGM"
	case 3
		strresearch = "resSF"
	case 4
		strresearch = "resMPV"
	case 5
		strresearch = "resBGP"
	case 6
		strresearch = "resMWP"
	case 7
		strresearch = "resMPC"
	case 8
		strresearch = "resWC"
	case 9
		strresearch = "resRP"
	case 10
		strresearch = "resUWR"
	case 11
		strresearch = "resFPC"
	case 12
		strresearch = "resEP"
	case 13
		strresearch = "resLB"
	case 14
		strresearch = "resG"
	case 15
		strresearch = "resSIT"
	case 16
		strresearch = "resBLC"
	case 17
		strresearch = "resSBS"
	case 18
		strresearch = "resMBS"
	case 19
		strresearch = "resAAS"
	case 20
		strresearch = "resFG"
	case 21
		strresearch = "resCS"
	case 22
		strresearch = "resWMP"
	case 23
		strresearch = "resE"
	case 24
		strresearch = "resD"
	case 25
		strresearch = "resAC"
end select

set rsResearch = server.createobject("ADODB.recordset")

StrSQL = "SELECT tblaccounts.username, food, wood, iron, money, researchticks, curResearch, resSF, resCE, resVGM, resBGP, resMPV, resMWP, resMPC, resFPC, resWC, resRP, resUWR, resG, resEP, resLB, resSIT, resBLC, resSBS, resMBS, resAAS, resAES, resCS, resFG, resWMP, resE, resD, resAC  FROM tblaccounts WHERE username = '" & strUsername & "';"

rsResearch.locktype = 2

rsResearch.Cursortype = 3

rsResearch.open StrSQL, adocon

if not rsResearch.eof then

select case strresearch
	case "resCE"
		check = true
		ticks = 10
		food = 500
		wood = 0
		iron = 1000
		money = 1250
	case "resVGM"
		if rsResearch("resCE") = 3 then
			check = true
			ticks = 12
			food = 500
			wood = 0
			iron = 850
			money = 1000
		end if
	case "resSF"
		check = true
		ticks = 8
		food = 500
		wood = 200
		iron = 500
		money = 2000
	case "resMPV"
		if rsResearch("resCE") = 3 then
			check = true
			ticks = 15
			food = 700
			wood = 500
			iron = 2000
			money = 2000
		end if
	case "resBGP"
		if rsResearch("resMPV") = 3 and rsResearch("resVGM") = 3 then
			check = true
			ticks = 10
			food = 700
			wood = 0
			iron = 1500
			money = 1750
		end if
	case "resMWP"
		if rsResearch("resCE") = 3 then
			check = true
			ticks = 18
			food = 500
			wood = 500
			iron = 1000
			money = 2500
		end if
	case "resMPC"
		if rsResearch("resMWP") = 3 then
			check = true
			ticks = 20
			food = 750
			wood = 0 
			iron = 1500
			money = 2500
		end if
	case "resWC"
		if rsResearch("resMWP") = 3 then
			check = true
			ticks = 15
			food = 1000
			wood = 0
			iron = 250
			money = 2500
		end if
	case "resRP"
		if rsResearch("resMWP") = 3 then
			check = true
			ticks = 24
			food = 1500
			wood = 100
			iron = 750
			money = 2750
		end if
	case "resUWR"
		if rsResearch("resRP") = 3 then
			check = true
			ticks = 18
			food = 1250
			wood = 0
			iron = 750
			money = 2500
		end if
	case "resFPC"
		if rsResearch("resMWP") = 3 then
			check = true
			ticks = 20
			food = 750
			wood = 1500
			iron = 0
			money = 2000
		end if
	case "resEP"
		if rsResearch("resMPC") = 3 then
			check = true
			ticks = 24
			food = 1500
			wood = 1000
			iron = 1000
			money = 1750
		end if
	case "resLB"
		if rsResearch("resEP") = 3 then
			check = true
			ticks = 16
			food = 1000
			wood = 500
			iron = 750
			money = 1250
		end if
	case "resG"
		if rsResearch("resLB") = 3 then
			check = true
			ticks = 36
			food = 1000
			wood = 250
			iron = 500
			money = 3750
		end if
	case "resSIT"
		check = true
		ticks = 8
		food = 750
		wood = 500
		iron = 500
		money = 1250
	case "resBLC"
		if rsResearch("resSIT") = 3 then
			check = true
			ticks = 12
			food = 1250
			wood = 1000
			iron = 1500
			money = 2000
		end if
	case "resSBS"
		check = true
		ticks = 48
		food = 3250
		wood = 2500
		iron = 5500
		money = 4750
	case "resMBS"
		if rsResearch("resSBS") = 3 then
			check = true
			ticks = 36
			food = 4500
			wood = 3000
			iron = 6750
			money = 5500
		end if
	case "resAAS"
		if rsResearch("resSBS") = 3 then
			check = true
			ticks = 65
			food = 6750
			wood = 3500
			iron = 8250
			money = 8250
		end if
	case "resFG"
		check = true
		ticks = 15
		food = 750
		wood = 500
		iron = 1500
		money = 1250
	case "resCS"
		check = true
		ticks = 18
		food = 1000
		wood = 0
		iron = 250
		money = 1250
	case "resWMP"
		check = true
		ticks = 18
		food = 750
		wood = 1250
		iron = 1750
		money = 1500
	case "resE"
		check = true
		ticks = 32
		food = 1750
		wood = 50
		iron = 150
		money = 5750
	case "resD"
		check = true
		ticks = 16
		food = 1750
		wood = 500
		iron = 1500
		money = 2500
	case "resAC"
		if rsResearch("resE") = 3 then
			check = true
			ticks = 20
			food = 2000
			wood = 1200
			iron = 1000
			money = 4000
		end if
end select
	if check = true and clng(food) =< clng(lngfood) and clng(wood) =< clng(lngwood) and clng(iron) =< clng(lngiron) and clng(money) =< clng(lngmoney) then
		rsResearch.fields("researchticks") = ticks
		rsResearch.fields("curResearch") = right(strresearch, cint(len(strresearch)-3))
		rsResearch.fields(strresearch) = 2
		rsresearch.fields("food") = clng(rsresearch("food")) - clng(food)
		rsresearch.fields("wood") = clng(rsresearch("wood")) - clng(wood)
		rsresearch.fields("iron") = clng(rsresearch("iron")) - clng(iron)
		rsresearch.fields("money") = clng(rsresearch("money")) - clng(money)
		rsResearch.update
	end if
	rsresearch.close
	set rsResearch = nothing
	
end if
response.redirect("research.asp")
end if
%>

<!---#include file="includes/end_check.asp"--->