<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
%>

<!---#include file="includes\top.asp"--->
<center>
<BR />
<BR />
<Table cellpadding="2" cellspacing="0" border="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline">
			From
		</TD>
		<TD class="tdheaderline">
			To
		</TD>
		<TD class="tdheaderline">
			Subject
		</TD>
		<TD class="tdheaderline">
			Date Sent
		</TD>
		<TD class="tdheaderline">
			Delete?
		</TD>
	</TR>
<%
Dim rsMessage

strSQL = "SELECT fromcountry, fromcontinent, tocontinent, tocountry, subject, datesent, messageid FROM tblMessages;"

set rsMessage = server.CreateObject("ADODB.Recordset")

rsMessage.Open strSQL, adocon

do while not rsMessage.EOF 
%>
	<TR>
		<TD class="norm">
			<% =rsMessage("fromcontinent") %>:<% =rsmessage("fromcountry") %>
		</TD>
		<TD class="norm">
			<% =rsMessage("tocontinent") %>:<% =rsmessage("tocountry") %>
		</TD>
		<TD class="norm">
			<A HREF="admin_view_msg2.asp?msg=<%=rsmessage("messageid") %>"><% =rsmessage("subject") %></A>
		</TD>
		<TD class="norm">
			<% =rsmessage("datesent") %>
		</TD>
		<TD class="norm">
			<A HREF="admin_del_msg.asp?msg=<% =rsmessage("messageid") %>">Delete</A>
		</TD>
	</TR>
<%
rsMessage.MoveNext 
loop

rsMessage.Close 

set rsmessage = nothing
%>

</table>
<BR />
<BR />

<!---#include file="includes\footer.asp"--->

<%
end if
%>
<!---#include file="includes\end_check.asp"--->