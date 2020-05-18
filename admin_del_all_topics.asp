<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else

	
	'code goes below here
Dim rsdelete
	
set rsdelete = server.createobject("ADODB.recordset")

StrSql = "SELECT tblcontopic.* FROM tblcontopic;"

rsdelete.locktype = 2

rsdelete.cursortype = 3

rsdelete.open StrSQL, adocon

do while not rsdelete.EOF 

	rsdelete.delete
	rsdelete.Movenext

loop
rsdelete.close

StrSql = "SELECT tblconpost.* FROM tblconpost;"

rsdelete.locktype = 2

rsdelete.cursortype = 3

rsdelete.open StrSQL, adocon

do while not rsdelete.eof
	rsdelete.delete
	rsdelete.movenext
loop

rsdelete.close

set rsdelete = nothing

response.redirect("admin.asp?msg=9)
	
	
		'code goes above here
	
end if
%>
<!---#include file="includes\end_check.asp"--->