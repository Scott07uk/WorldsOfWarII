<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else

'code goes below here

Dim rsConfig
strSQL = "SELECT AllowedSignups, ticksEnabled, AttacksAllowed, DefenceAllowed, ResAllowed, ConAllowed FROM tblconfig;"

set rsConfig = server.CreateObject("ADODB.Recordset")

rsConfig.LockType = 3

rsConfig.CursorType = 2

rsConfig.Open strsql, adocon

select case Request.QueryString("type")

case "ticks"
	if cbool(rsConfig("ticksEnabled")) = true then
			call rsConfig.Update ("ticksenabled", 0)
		else
			call rsConfig.Update ("ticksenabled", 1)
	end if
	
case "signup"
	if cbool(rsConfig("AllowedSignups")) = true then
			call rsConfig.Update ("AllowedSignups", 0)
		else
			call rsConfig.Update ("AllowedSignups", 1)
	end if
	
case "attacks"
		if cbool(rsConfig("attacksallowed")) = true then
			call rsConfig.Update ("attacksallowed", 0)
		else
			call rsConfig.Update ("attacksallowed", 1)
	end if
	
case "defence"
	if cbool(rsConfig("defenceallowed")) = true then
			call rsConfig.Update ("defenceallowed", 0)
		else
			call rsConfig.Update ("defenceallowed", 1)
	end if
	
case "research"
	if cbool(rsConfig("Resallowed")) = true then
			call rsConfig.Update ("resallowed", 0)
		else
			call rsConfig.Update ("resallowed", 1)
	end if
	
case "construction"
	if cbool(rsConfig("conallowed")) = true then
			call rsConfig.Update ("conallowed", 0)
		else
			call rsConfig.Update ("conallowed", 1)
	end if
end select

rsConfig.Close 

set rsconfig = nothing

Response.Redirect("admin.asp?msg=10")

'code goes above here
end if
%>
<!---#include file="includes\end_check.asp"--->