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

StrSQL = "SELECT tblcontopic.topicid, continent, subject FROM tblcontopic WHERE TopicID = " & request.Querystring("topicID") & "AND continent = " & lngboardid

rstopic.open strSQL, adocon

If not rstopic.EOF then

subject = rstopic("subject")

rstopic.close

StrSQL = "SELECT tblconpost.* FROM tblconpost WHERE TopicID = " & request.Querystring("topicID")

rstopic.open strSQL, adocon

If rstopic.EOF then
%>

<center>
<font color="#ff0000">There are no posts in this topic.</font>
</center>

<%
else
%>

<BR />
<table align="center" bgcolor="#0D1724" cellpadding="3" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<tr>
		<td colspan="2" class="tdheaderline">
			Topic: <%=subject%>
		</td>
	</tr>
<%
do while not rstopic.eof
%>

	<tr>
		<td width="20%" class="tdheaderline" style="border-right: solid; border-width:1;" NOWRAP>
			<%=rstopic("Postedby")%><br />
			<%=left(rstopic("dateposted"), len(rstopic("dateposted"))-3)%>
			<br />
			<a href="admin_board_Delete_post.asp?contid=<% =lngboardid %>&msgid=<%=rstopic("postID")%>&topicID=<%=request.querystring("topicID")%>"><font color="#ffff00">Delete</font></a>
		</td>
		<td width="100%" class="all" valign="top">
			<%=rstopic("message")%>
		</td>
	</tr>
<%
	rstopic.movenext
	loop
end if
end if

rstopic.close
set rstopic = nothing
%>
	<tr>
		<td colspan="3" align="right">
			<a href="admin_view_board2.asp?contid=<% =lngboardid %>"><font color="#ffff00">Back</font></a>
		</td>
	</tr>
</table>
<BR />

<!---#include file="includes\footer.asp"--->

<%
	end if
end if
%>
<!---#include file="includes\end_check.asp"--->