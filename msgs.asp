<!---#include file="includes/check.asp"--->
<!---#include file="includes/top.asp"--->

<!---SPELL CHECKED--->
<center>
<table width="95%" cellpadding="0" border="0" cellspacing="0" style="border-top:none; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD align="center" class="norm">
			<A HREF="javascript:openHelp('5');">Help With This Page</A>
		</TD>
	</TR>
</table>
</center>
<%
Dim rsmessagelist 'List messages in inbox.
Set rsmessagelist = Server.CreateObject("ADODB.Recordset")
%>

<Script language="javascript">
function selectAll() {
for (counter=0;counter<document.deletemsg.box.length;counter++)
{
document.deletemsg.box[counter].checked=document.deletemsg.selectall.checked;
}
}
</Script>
<br />
<center>
<%
StrSQL = "SELECT TblMessages.FromContinent, FromCountry, Subject, DateSent, New, MessageID FROM TblMessages WHERE ToContinent = " & LngContinentNo & " AND ToCountry = " & LngCountryNo
If Request.QueryString("sent") = 1 Then
%>
Message Sent
<%
End If
If Request.QueryString("error") = 4 Then
%>
Anti-Spam Filter Activated
<%
End If
%>
<form action="deletemsg_engine.asp" method="post" name="deletemsg">
<table style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;" cellpadding="5" cellspacing="5" width="95%">
<tr>
<td class="tdheader" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;" width="20"><input onclick="javascript:selectAll();" type="checkbox" name="selectall"></td>
<td class="tdheader" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;" width="100"><b>
<%
'If not ordered by co-ordinates then display as link.
If Request.QueryString("order") = 1 Then
StrSQL = strSQL & " ORDER BY FromContinent, FromCountry ASC;"
%>
Co-ords
<%
Else
%>
<a href="msgs.asp?order=1">Co-ords</a>
<%
End If
%>
</b></td>
<td class="tdheader" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;" width="200"><b>
<%
'If not ordered by date received then display as link.
If Request.QueryString("order") = 2 Then
StrSQL = strSQL & " ORDER BY DateSent ASC;"
%>
Date Received
<%
Else
%>
<a href="msgs.asp?order=2">Date Received</a>
<%
End If
%>
</b></td>
<td class="tdheader" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;"><b>
<%
'If not ordered by subject then display as link.
If Request.QueryString("order") = 3 Then
StrSQL = strSQL & " ORDER BY Subject ASC;"
%>
Subject
<%
Else
%>
<a href="msgs.asp?order=3">Subject</a>
<%
End If
%>
</b></td>
</tr>
<%
rsmessagelist.Open strSQL, Adocon
If rsmessagelist.EOF Then
%>
</table>
No Messages
<%
Else
Do While Not rsmessagelist.EOF
%>

<tr>
<td class="norm">
<input type="checkbox" name="box" class="form" value="<% = rsmessagelist("MessageID") %>">
</td>

<td class="norm">
<%
If CBool(rsmessagelist("New")) = True Then
	Response.Write("<font color=#00ff00>")
End If
%>
<% = rsmessagelist("FromContinent") %>:<% = rsmessagelist("FromCountry") %>
<%
If CBool(rsmessagelist("New")) = True Then
	Response.Write("</font>")
End If
%>
</td>

<td class="norm">
<%
If CBool(rsmessagelist("New")) = True Then
	Response.Write("<font color=#00ff00>")
End If
%>
<% = rsmessagelist("DateSent") %>
<%
If CBool(rsmessagelist("New")) = True Then
	Response.Write("</font>")
End If
%>
</td>

<td class="norm">
<%
If CBool(rsmessagelist("New")) = True Then
	Response.Write("<a href=""readmsg.asp?msg=" & rsmessagelist("MessageID") & """><font color=#00ff00>")
Else
	Response.Write("<a href=""readmsg.asp?msg=" & rsmessagelist("MessageID") & """>")
End If
%>
<% = rsmessagelist("Subject") %></a></td>
<%
If CBool(rsmessagelist("New")) = True Then
	Response.Write("</font>")
End If
%>
</tr>
<%
rsmessagelist.MoveNext
Loop
%>
</table>
<%
End If
rsmessagelist.Close
Set rsmessagelist = Nothing
%>
<br />
<br />
<table style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;" cellpadding="1" cellspacing="5" width="95%">
<td width="60"><input type="submit" value="Delete" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1; padding:0;" class="form"></td>
</form>
<td><form action="sendmsg.asp" method="post"></td>
<td width="60"><input type="submit" value="New Message" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1; padding:0" align="right" width="120" class="form"></td>
</form>
</table>
</center>

<!---#include file="includes/footer.asp"--->
<!---#include file="includes/end_check.asp"--->