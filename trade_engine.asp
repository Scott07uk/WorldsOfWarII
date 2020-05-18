<!---#include file="includes/check.asp"--->
<%
Dim lngValue		'the amount to change for
Dim rsExchange		'the object that does the work
Dim lngStart		'the starting ammount
Dim lngEnd			'the ending amount

'get the user input
lngValue = Request.Form("amount")

if len(lngValue) > 8 then
		lngValue="a"
end if

if blnTickRunning = true then
		Response.Redirect("development.asp?error=13")
	else

'validate the user input
if isnumeric(lngValue) = false then
		Response.Redirect("development.asp?error=8")
	else
		if lngValue < 1 then
				Response.Redirect("development.asp?error=12")
			else
		if request.form("from") = request.form("to") then
				response.redirect("development.asp?error=11")
			else
				set rsExchange = server.CreateObject("ADODB.Recordset")
		
				rsExchange.LockType = 3
				rsExchange.CursorType = 2
		
				strsql = "SELECT wood, iron, food, money FROM tblaccounts WHERE username = '" & strusername & "';"
		
				rsExchange.Open strsql, adocon
		
				if rsExchange.EOF then
						Response.Redirect("development.asp")
					else
						select case Request.Form("from")
							case 1
								lngstart = clng(rsexchange("food"))
							case 2
								lngstart = clng(rsexchange("wood"))
							case 3
								lngstart = clng(rsexchange("iron"))
							case 4
								lngstart = clng(rsexchange("money"))
						end select
				
						if lngstart < clng(lngValue) then
								Response.Redirect("development.asp?error=9")
							else
								select case Request.Form("from")
									case 1
										rsExchange.Fields("food") = clng(rsexchange("food")) - lngvalue
									case 2		
										rsExchange.Fields("wood") = clng(rsexchange("wood")) - lngvalue
									case 3
										rsExchange.Fields("iron") = clng(rsexchange("iron")) - lngvalue
									case 4
										rsExchange.Fields("money") = clng(rsexchange("money")) - lngvalue
								end select
						
						
								select case Request.Form("to")
									case 1
										rsExchange.Fields("food") = clng(rsexchange("food")) + (lngvalue * 0.75)
									case 2
										rsExchange.Fields("wood") = clng(rsexchange("wood")) + (lngvalue * 0.75)
									case 3
										rsExchange.Fields("iron") = clng(rsexchange("iron")) + (lngvalue * 0.75)
									case 4
										rsExchange.fields("money") = clng(rsexchange("money")) + (lngvalue *0.75)
								end select
							
								rsExchange.Update
							
								Response.Redirect("development.asp?error=10")
						end if
				end if
		
				rsExchange.Close 
		
				set rsexchange = nothing
		end if
end if
end if
end if

%>
<!---#include file="includes/end_check.asp"--->