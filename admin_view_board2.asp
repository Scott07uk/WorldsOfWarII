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
%>

<!---#include file="includes\top.asp"--->
<%
Dim rstopic
Set rstopic = Server.CreateObject("ADODB.Recordset")

StrSQL = "SELECT topicID, subject, startedby, lastpost, lastpostby FROM tblcontopic WHERE continent = " & lngBoardid & " ORDER BY lastpost DESC;"

rstopic.open strSQL, adocon

If rstopic.EOF then
%>

<center>
<font color="#ff0000">There are no topics to display.</font>
</center>

<%
else
%>
<center>
	<BR />
	Message board for continent <% =lngBoardid %>.
	<Br />
</center>
<BR />
<table align="center" bgcolor="#0D1724" cellpadding="3" cellspacing="0" width="95%" style="border-top:solid; border-bottom:none; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD width="55%" class="tdheaderline">
			Subject
		</td>
		<td width="15%" class="tdheaderline" NOWRAP>
			Started By
		</td>
		<td width="15%" class="tdheaderline" NOWRAP>
			Last Post By
		</td>
		<td width="15%" class="tdheaderline" NOWRAP>
			Date
		</td>
		<td class="tdheaderline">
			&nbsp;
		</td>
	</tr>
<%
do while not rstopic.eof
%>	
	<tr>
		<td width="50%" class="all" valign="top">
			<a href="admin_view_board_msg.asp?contid=<% =lngboardid %>&topicID=<%=rstopic("topicID")%>"><font color="#ffff00"><%=rstopic("subject")%></font></a>
		</td>
		<td width="10%" class="all" valign="top">
			<%=rstopic("Startedby")%>
		</td>
		<td width="10%" class="all" valign="top">
			<%=rstopic("lastpostby")%>
		</td>
		<td width="10%" class="all" valign="top">
			<%=left(rstopic("lastpost"), len(rstopic("lastpost"))-3)%>
		</td>
		<td class="all" valign="top">
			<a href="admin_del_board_topic.asp?contid=<% =lngboardid %>&topicID=<%=rstopic("topicID")%>"><font color="#ffff00">Delete</font></a>
		</td>
	</tr>
<%
	rstopic.movenext
	loop
end if

rstopic.close
set rstopic = nothing
%>
</table>
<Br />


<!---#include file="includes\footer.asp"--->

<%
	end if
end if
%>
<!---#include file="includes\end_check.asp"--->