<!---#include file="includes/check.asp"--->
<!---#include file="includes/top.asp"--->

<!---SPELL CHECKED--->
<center>
<table width="95%" cellpadding="0" border="0" cellspacing="0" style="border-top:none; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD align="center" class="norm">
			<A HREF="javascript:openHelp('6');">Help With This Page</A>
		</TD>
	</TR>
</table>
</center>
<br />
<center>
<%
If Request.QueryString("error") = 1 Then
	Response.Write("Empty")
End If

If Request.QueryString("error") = 2 Then
	Response.Write("Co-ordinates Must be Numeric")
End If

If Request.QueryString("error") = 3 Then
	Response.Write("The Co-ordinates Do Not Exist")
End If
%>
<form action="msg_engine.asp" method="post">
<table style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;" cellpadding="5" cellspacing="5" width="75%">
<tr>
<td align="right" class="tdheader">
CO-ORDS
</td>

<%

Dim strContinent
Dim strCountry
Dim strSubject
Dim strBody

If Request.QueryString("reply") = 1 Then
	Dim rsreplymessage 'Select previous message to reply to.
	StrSQL = "SELECT TblMessages.FromContinent, FromCountry, Subject, Body FROM TblMessages WHERE ToContinent = " & LngContinentNo & " AND ToCountry = " & LngCountryNo & " AND MessageID = " & Request.QueryString("Msg")
	Set rsreplymessage = Server.CreateObject("ADODB.Recordset")
	rsreplymessage.Open StrSQL, Adocon
End If

%>

<td class="norm">
<input type="text" name="continent" size="1" maxlength="4" class="form" value="
<%
If Request.QueryString("reply") = 1 Then
	Response.Write(rsreplymessage("FromContinent"))
End If
%>
">
<input type="text" name="country" size="1" maxlength="2" class="form" value="
<%
If Request.QueryString("reply") = 1 Then
	Response.Write(rsreplymessage("FromCountry"))
End If
%>
">
</td>
</tr>
<tr>
<td align="right" class="tdheader">
Subject
</td>
<td class="norm">
<input type="text" name="subject" size="65" maxlength="250" class="form" value="
<%
If Request.QueryString("reply") = 1 Then
	Response.Write("RE: " & rsreplymessage("Subject"))
End If
%>
">
</td>
</tr>
<tr>
<td align="right" valign="top" class="tdheader">
Message
</td>
<td class="norm"><textarea name="body" cols="60" rows="8" class="form"><%
If Request.QueryString("reply") = 1 Then
	Response.Write(vbCrlf)
	Response.Write(vbCrlf)
	Response.Write(vbCrlf)
	Response.Write("-----Original Message-----" & vbCrlf & vbCrlf)
	Dim strReply
	strReply = rsreplymessage("Body")

	strReply = Replace(strReply,"<BR>",vbCrlf)
	strReply = Replace(strReply,"<b>","[b]")
	strReply = Replace(strReply,"<b>","[B]")
	strReply = Replace(strReply,"</b>","[/b]")
	strReply = Replace(strReply,"</b>","[/B]")
	strReply = Replace(strReply,"<i>","[i]")
	strReply = Replace(strReply,"<i>","[I]")
	strReply = Replace(strReply,"</i>","[/i]")
	strReply = Replace(strReply,"</i>","[/I]")
	strReply = Replace(strReply,"<u>","[u]")
	strReply = Replace(strReply,"<u>","[U]")
	strReply = Replace(strReply,"</u>","[/u]")
	strReply = Replace(strReply,"</u>","[/U]")
	Response.Write(strReply)
End If
%></textarea></td>
</tr>
<tr>
<td></td>
<td>
<input type="submit" value="Send" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1; padding:0;" class="form">&nbsp;&nbsp;
<input type="reset" value="Reset" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1; padding:0;" class="form">
</form>
</td>
</tr>
</table>
<br />
<a href="msgs.asp">Inbox</a>
</center>

<!---#include file="includes/footer.asp"--->
<!---#include file="includes/end_check.asp"--->