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
if (document.frmSuspend.username.value == "") {if ((document.frmSuspend.country.value == "")&&(document.frmSuspend.continent.value == "")) {rtnVal=false;alert('You must enter either a username or location to edit.');}else{if ((document.frmSuspend.country.value == "")||(document.frmSuspend.continent.value == "")) {trnVal=false;alert('You have not entered a valid location to edit. ');}else{rtnVal=true;}}}
else{if ((document.frmSuspend.country.value == "")&&(document.frmSuspend.continent.value == "")) {rtnVal=true;}else{rtnval=false;alert('You must only enter either a username or a location to edit.');}}
return rtnVal;
}
//--->
</script>

<BR />
<BR />
<%
select case Request.QueryString("error")
	case 1
		Response.Write("You did not enter a username or a location to relocate. <BR /><BR />")
	case 2
		Response.Write("The location you entered was not valid.<BR /><BR />")
	case 3
		Response.Write("You entered both a location and a username. Enter only one.  <BR /><BR />")
	case 4
		Response.Write("The location you entered was not numeric.  <BR /><BR />")
	case 5
		Response.Write("The location you entered was out of range. Continents must be positive numbers and countries must be positive numbers between 1 and 15.  <BR /><BR />")
	case 6
		Response.Write("The account you wanted to move does not exist.  <BR /><BR />")
	case 7
		Response.Write("The location you entered for account relocation is occupied.  <BR /><BR />")
end select
%>
<center>
	<table cellpadding="2" border="0" cellspacing="0" width="95%" style="border-top:solid; border-right:solid; border-left:solid; border-bottom:solid; border-width:1; border-color:#FFFFFF;">
		<TR>
			<TD class="tdheaderline" align="center">
				Relocate Account
			</TD>
		</TR>
		<TR>
			<TD class="norm" align="center">
				Use this form to change a user's location. Do not use this feature unless it is absolutely necessary.
				<BR />
				<BR />
				<form action="admin_relocate_engine.asp" method="post" name="frmSuspend" onsubmit="return checkForm();">
					<table cellpadding="2" border="0" cellspacing="0">
						<TR>
							<TD class="tdheader" align="right">
								Username to Relocate
							</TD>
							<TD class="norm">
								<input type="text" name="username" size="20" maxlength="50" class="form" />
							</TD>
						</TR>
						<TR>
							<TD class="tdheader" align="right">
								OR Location to Relocate
							</TD>
							<TD class="norm">
								<input type="text" size="2" name="fromcontinent" maxlength="4" class="form" />
								:
								<input type="text" size="2" name="fromcountry" maxlength="2" class="form" />
							</TD>
						</TR>
						<TR>
							<TD class="tdheader" align="right">
								Relocate To
							</TD>
							<TD class="norm">
								<input type="text" size="2" name="tocontinent" maxlength="4" class="form" />
								:
								<input type="text" size="2" name="tocountry" maxlength="2" class="form" />
							</TD>
						</TR>
					</table>
					<BR />
					<input type="submit" value="Re-locate Account" class="form" />
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