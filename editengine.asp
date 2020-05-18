<!---#include file="includes/check.asp"--->

<!---SPELL CHECKED--->

<%
TopicID = request.querystring("topicID")
MsgID = request.querystring("msgID")

strMsg = trim( request.form("Message") )
strMsg = server.htmlencode(strMsg)
strMsg = Replace(strMsg,vbCrlf,"<BR>")

Set rstopic = Server.CreateObject("ADODB.Recordset")

StrSQL = "SELECT tblconpost.* FROM tblconpost WHERE postID = " & msgID

rstopic.locktype = 3

rstopic.cursortype = 2

rstopic.open strSQL, adocon

if not rstopic.eof then
	if rstopic("postedby") = strLeaderName then
		rstopic.fields("message") = strMsg
	end if
end if

rstopic.update
rstopic.close
set rstopic = nothing
adocon.close
set adocon = nothing

response.redirect("viewmsg.asp?topicID=" & topicid & "")
%>
<!---#include file="includes/end_check.asp"--->