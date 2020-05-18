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
Dim rsreadmessage 'List message to read.
StrSQL = "SELECT TblMessages.FromContinent, FromCountry, Subject, DateSent, New, Body, MessageID FROM TblMessages WHERE ToContinent = " & LngContinentNo & " AND ToCountry = " & LngCountryNo & " AND MessageID = " & Request.QueryString("msg")
Set rsreadmessage = Server.CreateObject("ADODB.Recordset")
rsreadmessage.LockType = 3
rsreadmessage.CursorType = 2
rsreadmessage.Open StrSQL, Adocon

If rsreadmessage.EOF Then
%>
No Message
<%
Else
Call rsreadmessage.Update("New", 0)
%>
<br />
<center>
<table style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;" cellpadding="5" cellspacing="5" width="75%">
<tr>
<td align="right" class="tdheader">
Co-ords
</td>
<td width="80%" class="norm"><% = rsreadmessage("FromContinent") %>:<% = rsreadmessage("FromCountry") %></td>
</td>
</tr>
<tr>
<td align="right" class="tdheader">
Date Received
</td>
<td class="norm"><% = rsreadmessage("DateSent") %></td>
</tr>
<tr>
<td align="right" class="tdheader">
Subject
</td>
<td class="norm"><% = rsreadmessage("Subject") %></a></td>
</tr>
<tr>
<td align="right" valign="top" class="tdheader">
Message
</td>
<td class="norm"><% = rsreadmessage("Body") %></td>
</tr>
</table>
<br />
<table border="0" cellspacing="6">
<form action="msgs.asp" method="post" name="inbox">
<input type="submit" value="Inbox" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1; padding:0;" class="form">
</form>
&nbsp;
<form action="deletemsg_engine.asp" method="post" name="deletemsg">
<input type="hidden" name="box" value="<% = Request.QueryString("msg")%>">
<input type="submit" value="Delete" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1; padding:0;" class="form">
</form>
&nbsp;
<form action="sendmsg.asp?reply=1&msg=<% = rsreadmessage("MessageID") %>" method="post" name="reply">
<input type="submit" value="Reply" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1; padding:0;" class="form">
</form>
</table>
</center>
<%
End If
rsreadmessage.Close
Set rsreadmessage = Nothing
%>

<!---#include file="includes/footer.asp"--->
<!---#include file="includes/end_check.asp"--->