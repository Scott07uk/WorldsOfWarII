<!---#include file="includes/check.asp"--->

<!---SPELL CHECKED--->

<%

'DECLAIR YOUR SODDING VARIBLES
'REMOVE THE SODING * KEY FROM YOUR KEYBOARD
'KILL YOUR SODING OBJECTS
'ASK ME TO SHOW YOU HOW TO WRITE THIS CODE IN ABOUT 20 LESS LINES WITH NO BUGGER UP THE DB TATICS LIKE YOU HAD
Dim rstopic
Dim intTopicNo
if request.form("subject") = "" then
	response.redirect("msgboard.asp?err=1")
elseif request.form("message") = "" then
	response.redirect("msgboard.asp?err=2")
else


strSubject = server.htmlencode(request.form("subject"))
strMsg = trim( request.form("Message") )
strMsg = server.htmlencode(strMsg)
strMsg = Replace(strMsg,vbCrlf,"<BR>")
 
set rsTopic = Server.CreateObject("ADODB.Recordset")

rstopic.LockType = 3

rstopic.CursorType = 2

strSQL = "SELECT * FROM tblcontopic ORDER BY topicid DESC;"

rstopic.Open strsql, adocon

if rstopic.EOF then
		inttopicno = 1 
	else
		intTopicNo = rsTopic("TopicID")
		intTopicNo = clng(intTopicNo) + 1
end if

rstopic.addnew

rstopic.fields("topicID") = inttopicNo
rstopic.fields("Startedby") = strLeaderName
rstopic.fields("replies") = 0
rstopic.fields("lastpost") = date() & " " & hour(now()) & ":" & minute(now())
rstopic.fields("lastpostby") = strLeaderName
rstopic.fields("subject") = strSubject
rstopic.fields("continent") = lngContinentno

rstopic.update
rstopic.close
set rstopic = nothing

Set rspost = Server.CreateObject("ADODB.Recordset")

rspost.cursortype=2

rspost.locktype=3

StrSQL = "SELECT tblconpost.* FROM tblconpost ORDER BY postID DESC;"

rspost.open strSQL, adocon

if rspost.eof then
		postID = 1
	else
		postID = rspost("postID")
		postID = clng(PostID) + 1
end if

rspost.addnew

rspost.fields("postID") = postID
rspost.fields("topicID") = inttopicno
rspost.fields("postedby") = strLeaderName
rspost.fields("dateposted") = Date() & " " & hour(now()) & ":" & minute(now())
rspost.fields("message") = strMsg

rspost.update
rspost.close
set rspost = nothing

response.redirect("msgboard.asp")
end if
%>
<!---#include file="includes/end_check.asp"--->