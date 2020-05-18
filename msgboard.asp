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
Dim blnpowerbase
Set rsuser = Server.CreateObject("ADODB.Recordset")

StrSQL = "SELECT powerbase FROM tblAccounts WHERE username = '" & strUsername & "';"

rsuser.open strSQL, adocon

If not rsuser.EOF then
	blnpowerbase = cbool(rsuser("powerbase"))
end if

rsuser.close
set rsuser = nothing

Set rstopic = Server.CreateObject("ADODB.Recordset")

StrSQL = "SELECT topicID, subject, startedby, lastpost, lastpostby FROM tblcontopic WHERE continent = " & lngContinentno & " ORDER BY lastpost DESC;"

rstopic.open strSQL, adocon

If rstopic.EOF then
%>

<center>
<font color="#ff0000">There are no topics to display</font>
</center>

<%
else
%>

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
		<td class="tdheaderline">&nbsp;
			
		</td>
	</tr>
<%
do while not rstopic.eof
%>	
	<tr>
		<td width="50%" class="all" valign="top">
			<a href="viewmsg.asp?topicID=<%=rstopic("topicID")%>"><font color="#ffff00"><%=rstopic("subject")%></font></a>
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
<% if rstopic("startedby") = strleadername then %>
		<td class="all" valign="top">
			<a href="deletetopic.asp?topicID=<%=rstopic("topicID")%>"><font color="#ffff00">Delete</font></a>
		</td>
<% else %>
		<td class="all" valign="top">&nbsp;
			
		</td>
<% end if %>
	</tr>
<%
	rstopic.movenext
	loop
end if

rstopic.close
set rstopic = nothing
%>
</table>

<BR />

<form action="newtopicengine.asp" method="post">

<table align="center" bgcolor="#0D1724" cellpadding="2" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<tr>
		<td class="tdheaderline" colspan="2" NOWRAP>
			New Topic
		</td>
	</tr>
<% 
if request.querystring("err") = 1 then 
%>
	<tr>
		<td class="all" colspan="2" align="center">
			<center><font color="#ff0000">The subject field is empty.</font></center>
		</td>
	</tr>
<% 
end if
if request.querystring("err") = 2 then 
%>
	<tr>
		<td class="all" colspan="2" align="center">
			<center><font color="#ff0000">The message field is empty.</font></center>
		</td>
	</tr>
<%
end if
%>
	<tr>
		<td class="all" style="border-right: solid; border-width: 1;">
			Subject
		</td>
		<td class="all">
			<input type="text" name="subject" size="50" class="form" valign="top">
		</td>
	</tr>
	<tr>
		<td class="all" style="border-right: solid; border-width: 1;" valign="top">
			Message
		</td>
		<td class="all">
			<textarea name="message" rows="5" cols="70" class="form"></textarea>
		</td>
	</tr>
	<tr>
		<td colspan="2" align="center" width="100%">
			<input type="submit" value="Create thread" class="form">
			<input type="reset" value="Reset fields" class="form">
		</td>
	</tr>
</table>
<!---#include file="includes/footer.asp"--->
<!---#include file="includes/end_check.asp"--->