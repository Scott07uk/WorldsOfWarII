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
if (document.frmSuspend.username.value == "") {if ((document.frmSuspend.country.value == "")&&(document.frmSuspend.continent.value == "")) {rtnVal=false;alert('You must enter either a username or location to suspend.');}else{if ((document.frmSuspend.country.value == "")||(document.frmSuspend.continent.value == "")) {trnVal=false;alert('You have not entered in a valid location to suspend. ');}else{rtnVal=true;}}}
else{if ((document.frmSuspend.country.value == "")&&(document.frmSuspend.continent.value == "")) {rtnVal=true;}else{rtnval=false;alert('You must only enter a username or a location to suspend.');}}
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
				<%
				select case Request.QueryString("error")
					case 1
						Response.Write("The account was not suspended because you did not enter a location or a username.  <BR /><BR />")
					case 2
						Response.Write("The account was not suspended because the location you entered was not numeric.  <BR /><BR />")
					case 3
						Response.Write("The account was not suspended because you entered both a location and a username.  <BR /><BR />")
					case 4
						Response.Write("You did not enter a reason for suspending the account.  <BR /><BR />")
					case 5
						Response.Write("There was no account matching that username or location.  <BR /><BR />")
				end select
				%>
				Use this form to suspend a users account. The account will remain on the system and can still be attacked, but the user will not be able to login. A reason for suspension must be provided for the email that will be sent to the user.
				<BR />
				<BR />
				<form action="admin_suspend_acc_engine.asp" method="post" name="frmSuspend" onsubmit="return checkForm();">
					<table cellpadding="2" border="0" cellspacing="0">
						<TR>
							<TD class="tdheader" align="right">
								Username to Suspend
							</TD>
							<TD class="norm">
								<input type="text" name="username" size="20" maxlength="50" class="form" />
							</TD>
						</TR>
						<TR>
							<TD class="tdheader" align="right">
								OR Location to Suspend
							</TD>
							<TD class="norm">
								<input type="text" size="2" name="continent" maxlength="4" class="form" />
								:
								<input type="text" size="2" name="country" maxlength="2" class="form" />
							</TD>
						</TR>
						<TR>
							<TD class="tdheader" align="right" valign="top">
								Reason for Suspension
							</TD>
							<TD class="norm">
								<textarea cols="50" rows="7" name="reason" class="form"></TEXTAREA>
							</TD>
						</TR>
					</table>
					<BR />
					<input type="submit" value="Suspend Account" class="form">
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