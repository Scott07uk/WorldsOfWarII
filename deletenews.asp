<!---#include file="includes/check.asp"--->

<!---SPELL CHECKED--->

<%
Dim rsNews

set rsNews = Server.CreateObject("ADODB.Recordset")

strSQL = "SELECT * FROM tblnews WHERE username = '" & strUsername & "';"

rsnews.LockType = 3

rsnews.CursorType = 2

rsNews.Open strSql, adocon

if not rsNews.EOF then
	do while not rsnews.EOF
		if not rsNews.EOF then
			rsnews.Delete
		end if
		rsnews.MoveNext
	loop
end if

rsnews.Close
set rsnews = nothing

Response.Redirect("priv_news.asp")
%>

<!---#include file="includes/end_check.asp"--->