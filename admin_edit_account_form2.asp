<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
Dim rsEdit			'recordset varible
Dim intcountry		'country number
Dim streditName		'the username of the nation being edited
Dim intcontinent	'the continent of the nation being editid
Dim intError		'error code
Dim blnUseLoc		'varible to decide if username or location is used

strEditName = Request.Form("username")
intContinent = Request.Form("continent")
intCountry = Request.Form("country")
interror = 0

'validate the information sumitted
if strSuspendName = "" then
		if (intCountry = "") or (intContinent = "") then
				intError=1
			else
				'we are using the location
				blnUseLoc = True
				'validate to make sure the numbers are correct
				if (isnumeric(intcountry)=False) or (isnumeric(intcontinent)=False) then
						interror = 2
					else
						'the submission is fine
				end if
		end if
	else
		if (intCountry <> "") or (intContinent <> "") then
				intError = 3
			else
				blnUseLoc=False
		end if
end if

if error <> 0 then
		Response.Redirect("admin_edit_account_form1.asp?error=" & interror)
	else
		if (isnumeric(intcontinent) = false) OR (isnumeric(intcounter) = false) then
				Response.Redirect("admin_edit_account_form2.asp?error=4")
			else

				strsql = "SELECT username, email, county, country, countryname, leadername, dob, lastlogin, holiday, continent, countryid, suspended, numholidays, admin, protectiontime, ticksmsg, firstlogin, IRCNick FROM tblaccounts WHERE "
		
				if blnUseLoc = True then
						strsql = strsql & "continent = " & intcontinent & " AND countryid = " & intcountry
					else
						strsql = strsql & "username = '" & streditname & "';"
				end if
				
				'create the object to retrieve the current settings
				
				set rsEdit = server.CreateObject("ADODB.Recordset")
				
				rsEdit.Open strsql, adocon
				
				if rsEdit.EOF then
						Response.Redirect("admin_edit_account_form1.asp?error=5")
					else
						


%>

<!---#include file="includes\top.asp"--->

<BR />
<BR />
<center>
	<table cellpadding="5" cellspacing="0" border="0" width="95%" style="border-top:solid;border-bottom:solid;border-left:solid;border-right:solid;border-width:1;border-color:#FFFFFF;">
		<TR>	
			<TD class="tdheaderline" align="center">
				Edit Account
			</TD>
		</TR>
		<TR>
			<TD class="norm" align="center">
				You can edit the account details below.
				<BR />
				<BR />
				<form action="admin_edit_account_engine.asp" method="post">
					<input type="hidden" name="username1" value="<% =rsEdit("username") %>" />
					<table cellpadding="2" cellspacing="0" border="0">
						<TR>
							<TD class="tdheader" align="right">
								Username
							</TD>
							<TD class="norm">
								<% =rsedit("username") %>
								<input type="hidden" name="username" value="<% =rsedit("username") %>" />
							</TD>
						</TR>
						<TR>
							<TD class="tdheader" align="right">
								Location
							</TD>
							<TD class="norm">
								<% =rsedit("continent") %>:<% =rsedit("countryid") %>
							</TD>
						</TR>
						<TR>
							<Td class="tdheader" align="right">
								Email Address
							</TD>
							<TD class="norm">
								<input type="text" size="40" name="email" class="form" maxlength="150" value="<% =rsedit("email") %>" />
							</TD>
						</TR>
						<TR>
							<TD class="tdheader" align="right">
								County
							</TD>
							<TD class="norm">
								<input type="text" size="15" name="county" maxlength="50" class="form" value="<% =rsedit("county") %>" />
							</TD>
						</TR>
						<TR>
							<TD class="tdheader" align="right">
								Country
							</TD>
							<TD class="norm">
								<input type="text" size="15" name="country" maxlength="50" class="form" value="<% =rsedit("country") %>" />
							</TD>
						</TR>
						<TR>
							<TD class="tdheader" align="right">
								Country Name
							</TD>
							<TD class="norm">
								<input type="text" size="30" name="countryname" maxlength="100" class="form" value="<% =rsedit("countryname") %>" />
							</TD>
						</TR>
						<TR>
							<TD class="tdheader" align="right">
								Leader Name
							</TD>
							<TD class="norm">
								<input type="text" size="30" name="leadername" maxlength="100" class="form" value="<% =rsedit("leadername") %>" />
							</TD>
						</TR>
						<TR>
							<TD class="tdheader" align="right">
								Date of Birth
							</TD>
							<TD class="norm">
								<% =rsedit("dob") %>
							</TD>
						</TR>
						<TR>
							<TD class="tdheader" align="right">
								Last Login
							</TD>
							<TD class="norm">
								<% =rsedit("lastlogin") %>
							</TD>
						</TR>
						<TR>
							<TD class="tdheader" align="right">
								On Holiday
							</TD>
							<TD class="norm">
								<%
								if cbool(rsedit("holiday")) = true then
								%>
								Yes
								<%
								else
								%>
								No
								<%
								end if
								%>
							</TD>
						</TR>
						<TR>
							<TD class="tdheader" align="right">
								Suspended
							</TD>
							<TD class="norm">
								<%
								if cbool(rsedit("suspended")) = true then
								%>
								Yes
								<%
								else
								%>
								No
								<%
								end if
								%>
							</TD>
						</TR>
						<TR>
							<TD class="tdheader" align="right">
								Number of Holidays
							</TD>
							<TD class="norm">
								<input type="text" size="2" class="form" name="numholidays" maxlength="2" value="<% =rsedit("numholidays") %>" />
							</TD>
						</TR>
						<TR>
							<TD class="tdheader" align="right">
								Administrator
							</TD>
							<TD class="norm">
								<%
								if cbool(rsedit("admin")) = true then
								%>
								Yes
								<%
								else
								%>
								No
								<%
								end if
								%>
							</TD>
						</TR>
						<TR>
							<TD class="tdheader" align="right">
								Protection Time
							</TD>
							<TD class="norm">
								<input type="text" size="2" name="protectiontime" maxlength="2" class="form" value="<% =rsedit("protectiontime") %>" />
							</TD>
						</TR>
						<TR>
							<TD class="tdheader" align="right">
								Messages Sent This Tick
							</TD>
							<TD class="norm">
								<input type="text" size="2" maxlength="2" name="ticksmsg" class="form" value="<% =rsedit("ticksmsg") %>" />
							</TD>
						</TR>
						<TR>
							<TD class="tdheader" align="right">
								User Ever Logged In?
							</TD>
							<TD class="norm">
								<%
								if cbool(rsedit("firstlogin")) = true then
								%>
								No
								<%
								else
								%>
								Yes
								<%
								end if
								%>
							</TD>
						</TR>
						<TR>
							<Td class="tdheader" align="right">
								IRC Nick
							</TD>
							<TD class="norm">
								<input type="text" size="15" name="ircnick" maxlength="25" class="form" value="<% =rsedit("IRCnick") %>" />
							</TD>
						</TR>
					</table>
					<BR />
					<input type="submit" value="Update Account" class="form" />
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
				rsEdit.Close 
				set rsedit = nothing
		end if
end if

end if
%>
<!---#include file="includes\end_check.asp"--->