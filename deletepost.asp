<!---#include file="includes/check.asp"--->

<!---SPELL CHECKED--->

<%
set rsaccount = server.createobject("ADODB.recordset")

strSql = "SELECT tblaccounts.username, powerbase FROM tblAccounts WHERE username = '" & strUsername & "';"

rsaccount.open StrSQL, adocon

if not rsaccount.eof then
	powerbase = rsaccount("powerbase")
end if
rsaccount.close
set rsaccount = nothing

set rsdelete = server.createobject("ADODB.recordset")

StrSql = "SELECT tblconpost.* FROM tblconpost WHERE postID = " & request.querystring("msgID") & "AND topicID = " & request.querystring("topicID")

rsdelete.locktype = 3

rsdelete.cursortype = 2

rsdelete.open StrSQL, adocon

if rsdelete("postedby") = strLeaderName then
	rsdelete.delete
elseif powerbase = true then
	rsdelete.delete
end if

rsdelete.close
set rsdelete = nothing

response.redirect("viewmsg.asp?topicID=" & request.querystring("topicID") & "")
%>

<!---#include file="includes/end_check.asp"--->
