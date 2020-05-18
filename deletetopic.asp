<!---#include file="includes/check.asp"--->

<%
topicID = Request.QueryString("topicID")

Set rsuser = Server.CreateObject("ADODB.Recordset")

StrSQL = "SELECT tblaccounts.username, powerbase FROM tblAccounts WHERE username = '" & strUsername & "';"

rsuser.open strSQL, adocon

If not rsuser.EOF then
	powerbase = cbool(rsuser("powerbase"))
end if

rsuser.close
set rsuser = nothing

set rsdelete = server.createobject("ADODB.recordset")

StrSql = "SELECT tblcontopic.* FROM tblcontopic WHERE topicID = " & topicID & "AND continent = " & lngContinentNo

rsdelete.locktype = 3

rsdelete.cursortype = 2

rsdelete.open StrSQL, adocon

if powerbase = false and rsdelete("startedby") <> strLeaderName then
	response.redirect("msgboard.asp")
else

	if not rsdelete.eof then

		rsdelete.delete

	end if
	rsdelete.close

	StrSql = "SELECT tblconpost.* FROM tblconpost WHERE topicID = " & topicID

	rsdelete.locktype = 3

	rsdelete.cursortype = 2

	rsdelete.open StrSQL, adocon
		
		do while not rsdelete.eof
			rsdelete.delete
			rsdelete.MoveNext
		loop

	rsdelete.close
	set rsdelete = nothing

	response.redirect("msgboard.asp?val=1")
end if
%>

<!---#include file="includes/end_check.asp"--->
