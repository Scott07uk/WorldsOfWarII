<!---#include file="includes/check.asp"--->

<!---SPELL CHECKED--->

<%
'declair varibles
Dim rsExplore			'recordset to do all the processing
Dim lngCost				'how many explorers will be needed
Dim intExploreLand		'how much land the user wants to explore
Dim intCounter
Dim lngInitialCost
Dim intgoldFound

intCounter = 1
intgoldfound = 0

'get the input from the user
intExploreLand = Request.Form("acres")

'validate the users input
intExploreLand = trim(intExploreLand)
if blnTickRunning = True then
		Response.Redirect("explore.asp?error=12")
	else
if isnumeric(intExploreLand) = false then
		Response.Redirect("explore.asp?error=8")
	else
		'format the number to ensure no decimal
		intExploreLand = cint(intExploreLand)
		if intExploreLand < 1 then
				Response.Redirect("explore.asp?error=9")
			else
				'calculate the cost
				'open the database to get the user current land level
				set rsExplore = server.CreateObject("ADODB.Recordset")
				
				strsql = "SELECT amountland, numexplorers, gold FROM tblaccounts WHERE username = '" & strusername & "';"
				
				rsExplore.LockType = 3
				
				rsExplore.CursorType = 2
				
				rsExplore.Open strsql, adocon
				
				lngInitialCost = cint(rsExplore("amountLand"))
				
				for intCounter = 1 to intExploreLand
					lngCost = clng(lngcost + lngInitialCost + intCounter -1)
					Randomize(second(now())*intCounter)
					if rnd() < 0.05 then
							intgoldFound = cint(intGoldFound + (rnd() * 5))
					end if
				next
				
				if clng(rsExplore("numexplorers")) < lngCost then
						Response.Redirect("explore.asp?error=10")
					else
						rsExplore.Fields("amountland") = clng(rsExplore("amountland")) + intExploreLand
						rsExplore.Fields("numexplorers") = clng(rsExplore("numexplorers")) - lngCost
						rsExplore.Fields("gold") = clng(rsExplore("gold")) + intgoldfound
						
						rsExplore.Update 
						
						Response.Redirect("explore.asp?error=11&gold=" & intgoldfound)
				end if
								
				rsExplore.Close 
				
				set rsexplore = nothing
		end if
end if
end if
%>
<!---#include file="includes/end_check.asp"--->