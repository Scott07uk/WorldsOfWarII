<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
%>

<!---#include file="includes\top.asp"--->

<%
Dim rsreadmessage 'List message to read.
StrSQL = "SELECT TblMessages.FromContinent, FromCountry, Subject, DateSent, New, Body, MessageID FROM TblMessages WHERE MessageID = " & Request.QueryString("msg")
Set rsreadmessage = Server.CreateObject("ADODB.Recordset")
rsreadmessage.Open StrSQL, Adocon

If rsreadmessage.EOF Then
%>
No Message
<%
Else
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
<BR />
<A HREF="admin_del_msg.asp?msg=<% =rsreadmessage("messageid") %>">Delete</A>
</center>
<%
End If
rsreadmessage.Close
Set rsreadmessage = Nothing
%>

<!---#include file="includes\footer.asp"--->

<%
end if
%>
<!---#include file="includes\end_check.asp"--->