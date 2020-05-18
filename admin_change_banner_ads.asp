<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
%>

<!---#include file="includes\top.asp"--->

<BR />
<BR />

<SCRIPT language="javascript">
//<-- Start hideing


function checkForm() {
var strUserName;
var rtnVal;

rtnVal = true;

strUserName = "<% =strusername %>"

if (document.frmBanner.username.value==strUserName){
	rtnVal =false;
	alert('You cannot edit your own account.');
}

return rtnVal;
}


//--->
</script>

<BR />
<BR />
<center>
	<%
	Select case Request.QueryString("error")
		case 1
			Response.Write("You did not enter a username to change.  <BR /><BR />")
		case 2
			Response.Write("The username you entered was invalid.  <BR /><BR />")
	end select
	%>
	<table cellpadding="2" border="0" cellspacing="0" width="95%" style="border-top:solid; border-right:solid; border-left:solid; border-bottom:solid; border-width:1; border-color:#FFFFFF;">
		<TR>
			<TD class="tdheaderline" align="center">
				Change Banner Advert Settings
			</TD>
		</TR>
		<TR>
			<TD class="norm" align="center">
				Use this page to turn the banner adverts on or off for an account.  
				<BR />
				<BR />
				<form name="frmBanner" action="admin_change_banner_engine.asp" method="post" onsubmit="return checkForm();">
					Username
					<input type="text" size="15" name="username" maxlength="50" class="form">
					<BR />
					<BR />
					Banner adverts should be <select name="setting" size="1" class="form">
						<option value="True">On</option>
						<option value="False" SELECTED>Off</option>
					</select>
					<BR />
					<BR />
					<input type="submit" value="Change" class="form">
				</form>
			</tD>
		</TR>
	</table>
</center>




<!---#include file="includes\footer.asp"--->

<%
end if
%>
<!---#include file="includes\end_check.asp"--->