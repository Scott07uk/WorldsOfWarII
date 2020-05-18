<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
'code goes below here

Dim rsmanual

strsql = "SELECT * FROM tblmanualcomments WHERE manid = " & Request.QueryString("ID")

set rsmanual = server.CreateObject("ADODB.Recordset")

rsmanual.LockType = 3
rsmanual.CursorType = 2

rsmanual.Open strsql, adocon

do while not rsmanual.EOF 
	rsmanual.Delete 
	rsmanual.Movenext
loop

rsmanual.Close 

set rsmanual = nothing

%>
<SCRIPT language="javascript">
window.close()
</script>
<%

'code goes above here
end if
%>
<!---#include file="includes\end_check.asp"--->