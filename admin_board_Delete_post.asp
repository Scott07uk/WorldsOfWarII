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

	StrSql = "SELECT tblconpost.* FROM tblconpost WHERE postID = " & request.querystring("msgID") & "AND topicID = " & request.querystring("topicID")

	rsdelete.locktype = 2

	rsdelete.cursortype = 3

	rsdelete.open StrSQL, adocon

	rsdelete.delete

	rsdelete.close
	set rsdelete = nothing

	response.redirect("admin_view_board_msg.asp?contid=" & lngboardid & "&topicID=" & request.querystring("topicID") & "")
	
	'code goes above here
	
	
	end if
end if
%>
<!---#include file="includes\end_check.asp"--->