<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
%>

<!---#include file="includes\top.asp"--->

<SCRIPT language="javascript">
//<---begin hiding

function checkForm(){
var rtnVal = false;
if (document.frmSuspend.username.value == "") {if ((document.frmSuspend.country.value == "")&&(document.frmSuspend.continent.value == "")) {rtnVal=false;alert('You must enter either a username or location to edit.');}else{if ((document.frmSuspend.country.value == "")||(document.frmSuspend.continent.value == "")) {trnVal=false;alert('You have not entered in a valid location to edit. ');}else{rtnVal=true;}}}
else{if ((document.frmSuspend.country.value == "")&&(document.frmSuspend.continent.value == "")) {rtnVal=true;}else{rtnval=false;alert('You must only enter a username or a location to edit.');}}
return rtnVal;
}
//--->
</script>
<BR />
<BR />
<center>
	<table cellpadding="2" border="0" cellspacing="0" width="95%" style="border-top:solid; border-right:solid; border-left:solid; border-bottom:solid; border-width:1; border-color:#FFFFFF;">
		<TR>
			<TD class="tdheaderline" align="center">
				Suspend Account
			</TD>
		</TR>
		<TR>
			<TD class="norm" align="center">
				Use this form to change the settings for an account.  
				<BR />
				<BR />
				<form action="admin_edit_account_form2.asp" method="post" name="frmSuspend" onsubmit="return checkForm();">
					<table cellpadding="2" border="0" cellspacing="0">
						<TR>
							<TD class="tdheader" align="right">
								Username Edit
							</TD>
							<TD class="norm">
								<input type="text" name="username" size="20" maxlength="50" class="form" />
							</TD>
						</TR>
						<TR>
							<TD class="tdheader" align="right">
								OR Location to Edit
							</TD>
							<TD class="norm">
								<input type="text" size="2" name="continent" maxlength="4" class="form" />
								:
								<input type="text" size="2" name="country" maxlength="2" class="form" />
							</TD>
						</TR>
					</table>
					<BR />
					<input type="submit" value="Edit Account" class="form">
				</form>
				<BR />
				<BR />
				<A Href="admin.asp">Main Admin Menu</A>
			</TD>
		</TR>
	</table>
</center>

<!---#include file="includes\footer.asp"--->

<%
end if
%>
<!---#include file="includes\end_check.asp"--->