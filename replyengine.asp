<!---#include file="includes/check.asp"--->

<!---SPELL CHECKED--->

<%
Dim ID
Dim strPos
Dim strMsg
ID = request.form("ID")

if request.form("message") = "" then
	response.redirect("viewmsg.asp?topicID=" & ID & "&err=1")
end if

strMsg = trim( request.form("Message") )
strMsg = server.htmlencode(strMsg)
strMsg = Replace(strMsg,vbCrlf,"<BR>")
strMsg = Replace(strMsg,"[b]","<B>")
strMsg = Replace(strMsg,"[/b]","</b>")
strMsg = Replace(strMsg,"[u]","<u>")
strMsg = Replace(strMsg,"[/u]","</u>")
strMsg = replace(strMsg,"[i]","<i>")
strMsg = replace(strMsg,"[/i]","</i>")

Set rstopic = Server.CreateObject("ADODB.Recordset")

rstopic.locktype=3

rstopic.cursortype=2

StrSQL = "SELECT tblconpost.* FROM tblconpost ORDER BY postID DESC;"

rstopic.open strSQL, adocon

if rstopic.eof then
		postID = 1
	else
		postID = rstopic("postID")
		postID = clng(PostID) + 1
end if

rstopic.addnew

rstopic.fields("postID") = postID
rstopic.fields("TopicID") = ID
rstopic.fields("postedby") = strLeaderName
rstopic.fields("Dateposted") = date() & " " & hour(now()) & ":" & minute(now())
rstopic.fields("message") = strMsg

rstopic.update
rstopic.close

StrSQL = "SELECT lastpostby, replies, lastpost FROM tblcontopic WHERE topicID = " & ID

rstopic.locktype=3

rstopic.cursortype=2

rstopic.open strSQL, adocon

rstopic.fields("lastpostby") = strLeadername
rstopic.fields("replies") = rstopic("replies") + 1
rstopic.fields("lastpost") = date() & " " & hour(now()) & ":" & minute(now())

rstopic.update
rstopic.close
set rstopic = nothing

response.redirect("viewmsg.asp?topicID=" & id & "")
%>

<!---#include file="includes/end_check.asp"--->