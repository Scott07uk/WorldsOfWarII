<!---#include file="includes/check.asp"--->
<!---#include file="includes/top.asp"--->

<!---SPELL CHECKED--->
<center>
<table width="95%" cellpadding="0" border="0" cellspacing="0" style="border-top:none; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD align="center" class="norm">
			<A HREF="javascript:openHelp('2');">Help With This Page</A>
		</TD>
	</TR>
</table>
</center>
<%
'DEANO THE ORDER THAT STUFF COMES OUT THE DB IS NOT AN APPROPRITE ORDERING, ESPICIALLY WHEN ITS LIKLEY TO BE IN THE WRONG ORDER ANY WAY
'GOD MADE THE ORDER BY STATEMENT FOR A SODDING REASON
set rsNews = Server.CreateObject("ADODB.Recordset")

strSQL = "SELECT tblnews.timemade, unread, message FROM tblnews WHERE username = '" & strUsername & "' ORDER BY timemade DESC;"

rsnews.LockType = 3

rsnews.Cursortype = 2

rsNews.Open strSql, adocon

if rsnews.EOF then
%>
<br />
<center>
	<font color="#ff0000">There is no news to view.</font>
</center>
<%
else
%>

<BR />
<table align="center" bgcolor="#0D1724" cellpadding="3" cellspacing="0" width="95%" style="border-top:solid; border-bottom:none; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<tr>
		<td class="tdheaderline" align="center" colspan="2">
			Private News
		</td>
	</tr>
<%
	do while not rsnews.EOF 
		rsnews.Fields("unread") = 0
		rsnews.update
%>
	<tr>
		<td width="7%" class="tdheaderline" style="border-right: solid; border-width: 1px;">
			<%=rsnews("timemade")%>
		</td>
		<td class="norm" valign="top" style="border-bottom:solid; border-width: 1px;">
			<%=rsnews("message")%>
		</td>
	</tr>
<%
	rsnews.MoveNext
	loop
end if

rsnews.Close

set rsnews = nothing
%>
<%
If blnTickRunning = False then
%>
	<tr>
		<td align="center" colspan="2" style="border-bottom: solid; border-width: 1px;">
			<a href="deletenews.asp">Delete News</a>
		</td>
	</tr>
	<%
	end if
	%>
</table>
<BR />
<!---#include file="includes/footer.asp"--->
<!---#include file="includes/end_check.asp"--->