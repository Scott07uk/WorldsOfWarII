<%
option explicit
%>
<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else

'code goes below here

Dim rsEmail

set rsemail = server.CreateObject("ADODB.Recordset")
				
rsEmail.LockType = 3
				
rsEmail.Cursortype = 2
				
strsql = "SELECT * FROM tblactiveemails;"
				
rsEmail.Open strsql, adocon
				
do while not rsEmail.EOF 
	rsEmail.Delete 
	rsEmail.MoveNext
loop

rsEmail.Close 

set rsemail = nothing

Response.Redirect("admin.asp?msg=15") 		

'code goes above here
end if
%>
<!---#include file="includes\end_check.asp"--->