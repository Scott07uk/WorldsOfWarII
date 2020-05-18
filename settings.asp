
<!---#include file="includes\check.asp"--->
<!---#include file="includes\top.asp"--->

<!---SPELL CHECKED--->
<center>
<table width="95%" cellpadding="0" border="0" cellspacing="0" style="border-top:none; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD align="center" class="norm">
			<A HREF="javascript:openHelp('19');">Help With This Page</A>
		</TD>
	</TR>
</table>
</center>

<script language="javascript">
//<---
function checkForm(){
var rtnVal;
if (document.settings.newpassword.value==document.settings.confirmnew.value){
	rtnVal=true;
}else{
	rtnVal=false;
	alert('The new passwords do not match.');
}
return rtnVal;
}
//--->
</script>

<SCRIPT language="javascript">
	//<--hide script from older browsers
	//function to open a new window		
	function openWin(url)
	{
		//open the new window
		window.open(url,'email', config='height=450,width=650,toolbar=no');
	}				
	//--end hiding>
</script>
<center>
<br>
<%
If Request.QueryString("error") = 1 Then
%>
The old password you entered was not correct.
<br>
<%
End If

If Request.QueryString("error") = 2 Then
%>
The new passwords you entered did not match.
<br>
<%
End If

If Request.QueryString("error") = 3 Then
%>
You may not go on holiday again this round.
<br>
<%
End If

If Request.QueryString("error") = 4 Then
%>
The email address you entered has not yet been activated.
Click on the activate link next to your email address.
<br>
<%
End If

If Request.QueryString("error") = 5 Then
%>
You did not enter a valid date of birth.
Leave the field blank if you wish.
<br>
<%
End If

If Request.QueryString("msg") = 1 Then
%>
Your settings have been updated.
<br>
<%
End If
%>
<br>
<form action="settings_engine.asp" method="post" name="settings" onsubmit="return checkForm();">
<Table cellpadding="2" cellspacing="0" border="0" width="95%">

	<tr>
		<td width="20%" class="tdheader" align="center" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			Change Password
		</td>

		<td valign="top" class="norm" align="left" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
		<table cellpadding="0" cellspacing="0">
			<tr>
				<td>
					Old
				</td>
			</tr>
			<tr>
				<td>
					<input type="password" name="oldpassword" size="20" maxlength="50" class="form">
				</td>
			</tr>
			<tr>
				<td>
					New
				</td>
			</tr>
			<tr>
				<td>
					<input type="password" name="newpassword" size="20" maxlength="50" class="form">
				</td>
			</tr>
			<tr>
				<td>
					Confirm New
				</td>
			</tr>
			<tr>
				<td>
					<input type="password" name="confirmnew" size="20" maxlength="50" class="form">
				</td>
			</tr>
		</table>
		</td>
	</tr>

<%
Dim rssettings 'Open account settings details.
Set rssettings = Server.CreateObject("ADODB.Recordset")
StrSQL = "SELECT TblAccounts.Email, County, Country, Dob, Holiday, NumHolidays, IRCNick FROM TblAccounts WHERE Continent = " & LngContinentNo & " AND CountryID = " & LngCountryNo
rssettings.Open strSQL, Adocon
%>

<tr>
<td width="20%" class="tdheader" align="center" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
Email
</td>

<td class="norm" align="left" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
<input type="text" name="email" size="40" maxlength="100" class="form" value="<% = rssettings("Email") %>"> &nbsp; <a href="javascript:openWin('activate_email1.asp');">Activate This Address</a>

</td>
</tr>

<tr>
<td width="20%" class="tdheader" align="center" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
County
</td>

<td class="norm" align="left" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
<input type="text" name="county" size="40" maxlength="150" class="form" value="<% = rssettings("County") %>">
</td>
</tr>

<tr>
<td width="20%" class="tdheader" align="center" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
Country
</td>

<td class="norm" align="left" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
<input type="text" name="country" size="40" maxlength="150" class="form" value="<% = rssettings("Country") %>">
</td>
</tr>

<tr>
<td width="20%" class="tdheader" align="center" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
Date of Birth
</td>

<td class="norm" align="left" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
<input type="text" name="dateofbirth" size="20" maxlength="50" class="form" value="<% = rssettings("Dob") %>">
</td>
</tr>

<tr>
<td width="20%" class="tdheader" align="center" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
Go on Holiday
</td>

<td class="norm" align="left" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
<input type="checkbox" name="holiday" value="yes" class="form">
<%
If cbool(rssettings.Fields("Holiday")) = False Then
%>
You ARE NOT on holiday.
<%
End If
If cbool(rssettings.Fields("Holiday")) = True Then
%>
You ARE on holiday.
<%
End If
%>
You have used <% = rssettings("NumHolidays") %> holidays.
</td>
</tr>

<tr>
<td width="20%" class="tdheader" align="center" style="border-top:solid; border-left:solid; border-right:solid; border-bottom:solid; border-width:1; border-color:#FFFFFF;">
IRC Nickname
</td>

<td class="norm" align="left" style="border-top:solid; border-left:solid; border-right:solid; border-bottom:solid; border-width:1; border-color:#FFFFFF;">
<input type="text" name="ircnick" size="20" maxlength="50" class="form" value="<% = rssettings("IRCNick") %>">
</td>
</tr>

<tr>
<td width="20%">
</td>

<td class="norm" align="left" style="border-top:solid; border-left:solid; border-right:solid; border-bottom:solid; border-width:1; border-color:#FFFFFF;">
<input type="submit" value="Save Changes" class="form">
</td>
</tr>

</table>
</form>

</center>
<!---#include file="includes\end_check.asp"--->
<!---#include file="includes\footer.asp"--->