<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
%>
<!---#include file="includes\top.asp"--->
<%
Dim strLongMessage
Dim strdbfield
Dim rsMsg
select case Request.QueryString("type")
	case "creator"
		strLongMessage = "message from creators"
		strdbfield = "creatorsmessage"
	case "register"
		strlongmessage = "registration email"
		strdbfield = "signupmsg"
end select

strsql = "SELECT " & strdbfield & " FROM tblconfig;"

set rsmsg = server.CreateObject("ADODB.Recordset")

rsMsg.Open strsql, adocon

%>
<BR />
<BR />
<center>
<table cellpadding="2" width="95%" border="0" cellspacing="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline" align="center">
			Change the <% =strlongmessage %>
		</TD>
	</TR>
	<TR>
		<TD class="norm" align="center">
			Use this form to change the <% =strlongmessage %>.
			<BR />
			<BR />
			<form action="admin_change_msgs_engine.asp" method="post">
				<input type="hidden" name="type" value="<% =request.querystring("type") %>" />
				<TEXTAREA name="message" cols="100" rows="12" class="form"><% =rsmsg(strdbfield) %></textarea>
				<Br />
				<BR />
				<input type="submit" value="Update" class="form" />
			</form>
			<BR />
			<BR />
			<A Href="admin.asp">Main Admin Menu</A>
		</TR>
	</TR>
</table>
</center>

<!---#include file="includes\footer.asp"--->

<%
end if
%>
<!---#include file="includes\end_check.asp"--->