<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
%>
<!---#include file="includes\top.asp"--->
<%
dim rsconfig

strsql = "SELECT gameaddress, portaladdress, forumsaddress, manualaddress, ircsite FROM tblconfig;"

set rsconfig = server.CreateObject("ADODB.Recordset")

rsconfig.Open strsql, adocon
%>
<script language="javascript">
//<---

function checkForm(){
var rtnVal;

if (document.frmemail.domain.value==""){
	if (document.frmemail.email.value==""){
		rtnVal=false;
		alert('You must enter either an email address or an email domain.');
	}else{
		rtnVal=true;
	}
}else{
	if (document.frmemail.email.value==""){
		rtnVal=true;
	}else{
		rtnVal=false;
		alert('You can only enter an email or a doamin, not both.');
	}
}
return rtnVal;
}

//--->
</script>
<BR />
<BR />
<center>
<%
select case Request.QueryString("error")
	case 1
		Response.Write("You did not enter either an email address or a domain to ban.  <BR /><BR />")
	case 2
		Response.Write("You can not enter both an email address and a domain to ban.  <BR /><BR />")
end select
%>

<table cellpadding="2" width="95%" border="0" cellspacing="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline" align="center">
			Add Banned Email
		</TD>
	</TR>
	<TR>
		<TD class="norm" align="center">
			You can use this page to ban email addresses or entire email domains from being used on the game.  
			<BR />
			<BR />
			<form action="admin_ban_email_engine.asp" method="post" name="frmemail" onsubmit="return checkForm();">
				<table cellpadding="2" border="0" cellpadding="0">
					<TR>
						<TD class="tdheader" align="right">
							Domain to Ban
						</TD>
						<TD>
							<input type="text" size="15" maxlength="50" class="form" name="domain" />
						</TD>
					</TR>
					<TR>
						<TD class="tdheader" align="right">
							OR Email to Ban
						</TD>
						<TD class="norm">
							<input type="text" size="35" maxlength="150" class="form" name="email" />
						</TD>
					</TR>
				</table>
				<input type="submit" value="Update" class="form">
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