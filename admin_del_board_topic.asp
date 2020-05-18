<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
Dim lngBoardid
lngBoardid = Request.QueryString("contid")
if isnumeric(lngboardid) = false then
		Response.Redirect("admin.asp")
	else
	
	'code goes below here
	Dim rsdelete
	
	set rsdelete = server.createobject("ADODB.recordset")

StrSql = "SELECT tblcontopic.* FROM tblcontopic WHERE topicID = " & request.querystring("topicID") & "AND continent = " & lngboardid

rsdelete.locktype = 2

rsdelete.cursortype = 3

rsdelete.open StrSQL, adocon

if not rsdelete.eof then

	rsdelete.delete

end if
rsdelete.close

StrSql = "SELECT tblconpost.* FROM tblconpost WHERE topicID = " & request.querystring("topicID")

rsdelete.locktype = 2

rsdelete.cursortype = 3

rsdelete.open StrSQL, adocon

do while not rsdelete.eof
	rsdelete.delete
	rsdelete.movenext
loop

rsdelete.close

set rsdelete = nothing

response.redirect("admin_view_board2.asp?contid=" & lngboardid)
	
	
		'code goes above here
	
	
	end if
end if
%>
<!---#include file="includes\end_check.asp"--->