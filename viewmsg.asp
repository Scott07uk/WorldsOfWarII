<!---#include file="includes/check.asp"--->
<!---#include file="includes/top.asp"--->

<!---SPELL CHECKED--->
<center>
<table width="95%" cellpadding="0" border="0" cellspacing="0" style="border-top:none; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD align="center" class="norm">
			<A HREF="javascript:openHelp('7');">Help With This Page</A>
		</TD>
	</TR>
</table>
</center>

<%
Dim blnPowerBase
Dim rsUser
Set rsUser = Server.CreateObject("ADODB.Recordset")

StrSQL = "SELECT powerbase FROM tblAccounts WHERE username = '" & strUsername & "';"

rsuser.open strSQL, adocon

If not rsuser.EOF then
	blnpowerbase = cbool(rsuser("powerbase"))
end if

rsuser.close
set rsuser = nothing

Set rstopic = Server.CreateObject("ADODB.Recordset")

StrSQL = "SELECT tblcontopic.topicid, continent, subject FROM tblcontopic WHERE TopicID = " & request.Querystring("topicID") & "AND continent = " & lngContinentNo

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
		<td colspan="2" class="tdheaderline" align="center">
			Topic: <%=subject%>
		</td>
	</tr>
<% if Request.QueryString("err") = 1 then %>
	<tr>
		<td colspan="2" align="center" style="border-bottom: solid; border-width: 1;">
			<font color="#ff0000">You did not enter a reply.</font>
		</td>
	</tr>
<%
end if
do while not rstopic.eof
%>

	<tr>
		<td width="135" class="tdheaderline" style="border-right: solid; border-width:1;" NOWRAP valign="top">
			<%=rstopic("Postedby")%><br />
			<%=left(rstopic("dateposted"), len(rstopic("dateposted"))-3)%>
<% if rstopic("postedby") = strLeaderName then %>
			<br />
			<a href="Editpost.asp?msgid=<%=rstopic("postID")%>&topicID=<%=request.querystring("topicID")%>"><font color="#ffff00">Edit</font></a>
			&nbsp;
			<a href="Deletepost.asp?msgid=<%=rstopic("postID")%>&topicID=<%=request.querystring("topicID")%>"><font color="#ffff00">Delete</font></a>
<% elseif blnpowerbase = true then %>
			<br />
			<a href="Deletepost.asp?msgid=<%=rstopic("postID")%>&topicID=<%=request.querystring("topicID")%>"><font color="#ffff00">Delete</font></a>
<% end if %>
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
			<a href="msgboard.asp"><font color="#ffff00">Back</font></a>
		</td>
	</tr>
</table>

<br />
<form action="replyengine.asp" method="post">
<input type="hidden" name="id" value="<%=request.querystring("topicID")%>" />
<table align="center" bgcolor="#0D1724" cellpadding="3" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<tr>
		<td class="tdheaderline">
			Reply
		</td>
	</tr>
<% if request.querystring("err") = 1 then %>
	<tr>
		<td class="all" align="center">
			<font color="#ff0000">The reply field was empty.</font>
		</td>
	</tr>
<% end if %>
	<tr>
		<td class="all">
			<textarea name="message" rows="7" cols="70" class="form"></textarea>
		</td>
	</tr>
	<tr>
		<td align="center">
			<input type="submit" value="Reply" class="form" />
			&nbsp;
			<input type="reset" value="Reset" class="form" />
		</td>
	</tr>
</table>
</form>
<!---#include file="includes/footer.asp"--->
<!---#include file="includes/end_check.asp"--->