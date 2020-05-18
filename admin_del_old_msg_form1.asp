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
<%
select case Request.QueryString("error")
	case 1
		Response.Write("The date you entered was not valid.  <BR /><BR />")
end select
%>
<table cellpadding="2" cellspacing="0" border="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline" align="center">
			Delete Old Messages
		</TD>
	</TR>
	<TR>
		<TD class="norm" align="center">
			Delete ALL Messages Older Than
			<BR />
			<form action="admin_del_msgs_engine.asp" method="post">
				<input type="text" size="10" name="deldate" maxlength="10" class="form" value="<% =cdate(now()-28) %>" />
				<BR />
				<BR />
				<input type="submit" value="Delete Messages" class="form" />
			</form>
		</TD>
	</TR>
</table>
<BR />
<BR />
</center>

<!---#include file="includes\footer.asp"--->

<%
end if
%>
<!---#include file="includes\end_check.asp"--->