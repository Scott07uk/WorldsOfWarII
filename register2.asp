<!---SPELL CHECKED--->

<!---#include file="includes/ptop.asp"--->
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
<table cellpadding="5" border="0" cellspacing="5" style="border-left:solid; border-right:solid; border-top:solid; border-bottom:solid; border-width:1; border-color:#FFFFFF;" width="95%">
	<TR>
		<TD class="tdheader" align="center">
			Register for Worlds of War II
		</TD>
	</TR>
	<TR>
		<TD class="norm" align="center">
			<form action="register3.asp" method="post">
				Please fill in the form below to create your account for Worlds of War II. Be sure to read the terms of service first.
				<BR />
				<BR />
				<%
				select case Request.QueryString("error")
				
				case 1
					Response.Write("You have not entered a username to register with.<BR /><BR />")
				case 2
					Response.Write("You have not entered an email to register with.<BR /><BR />")
				case 3
					Response.Write("You have not entered a name for your nation.<BR /><BR />")
				case 4
					Response.Write("You have not entered a leader for your nation.<BR /><BR />")
				case 5
					Response.Write("The two passwords were not the same.<BR /><BR />")
				case 6
					Response.Write("The date of birth is not valid.  <BR /><BR />")
				case 7
					Response.Write("The email address has not yet been activated in our system. Click <A HREF=""javascript:openWin('activate_email1.asp');"">here</A> to activate your email address.<BR /><BR />  ")
				case 8
					Response.Write("The email address you entered has been entered into our system, but you have not confirmed it is valid. Click <A HREF=""javascript:openWin('add_email2.asp');"">here</A> to confirm that it's active and working.<BR /><BR />")
				case 9
					Response.Write("Sorry but that username is already taken. Please choose another one.<BR /><BR />")
				end select
				%>
				<table cellpadding="2" cellspacing="0" border="0" style="border-left:solid; border-right:solid; border-top:solid; border-bottom:solid; border-width:1; border-color:#FFFFFF;">
					<TR>
						<TD class="tdheader" align="right">
							Username
						</TD>
						<TD class="norm">
							<input type="text" size="20" maxlength="50" name="username" value="<% =Request.QueryString("username") %>" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="tdheader" align="right">
							Password
						</TD>
						<TD class="norm">
							<input type="password" size="20" maxlength="20" name="pass1" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="tdheader" align="right">
							Confirm Password
						</TD>
						<TD class="norm">
							<input type="password" size="20" maxlength="50" name="pass2" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="tdheader" align="right">
							Email
						</TD>
						<TD class="norm">
							<input type="text" size="30" name="email" maxlength="150" class="form" value="<% =Request.QueryString("email") %>" />
						</TD>
					</TR>
					<TR>
						<TD class="tdheader" align="right">
							County
						</TD>
						<TD class="norm">
							<input type="text" size="15" name="county" maxlength="20" class="form" value="<% =Request.QueryString("fromcounty") %>" />
						</TD>
					</TR>
					<TR>
						<TD class="tdheader" align="right">
							Country
						</TD>
						<TD class="norm">
							<input type="text" size="15" name="country" maxlength="50" class="form" value="<% =Request.QueryString("fromcountry") %>" />
						</TD>
					</TR>
					<TR>
						<TD class="tdheader" align="right">
							Date of Birth
						</TD>
						<TD class="norm">
							<SELECT name="dobday" size="1" class="form">
								<option value="1">1</option>
								<option value="2">2</option>
								<option value="3">3</option>
								<option value="4">4</option>
								<option value="5">5</option>
								<option value="6">6</option>
								<option value="7">7</option>
								<option value="8">8</option>
								<option value="9">9</option>
								<option value="10">10</option>
								<option value="11">11</option>
								<option value="12">12</option>
								<option value="13">13</option>
								<option value="14">14</option>
								<option value="15">15</option>
								<option value="16">16</option>
								<option value="17">17</option>
								<option value="18">18</option>
								<option value="19">19</option>
								<option value="20">20</option>
								<option value="21">21</option>
								<option value="22">22</option>
								<option value="23">23</option>
								<option value="24">24</option>
								<option value="25">25</option>
								<option value="26">26</option>
								<option value="27">27</option>
								<option value="28">28</option>
								<option value="29">29</option>
								<option value="30">30</option>
								<option value="31">31</option>
							</SELECT>
							<SELECT name="dobmonth" size="1" class="form">
								<option value="1">January</option>
								<option value="2">February</option>
								<option value="3">March</option>
								<option value="4">April</option>
								<option value="5">May</option>
								<option value="6">June</option>
								<option value="7">July</option>
								<option value="8">August</option>
								<option value="9">September</option>
								<option value="10">October</option>
								<option value="11">November</option>
								<option value="12">December</option>
							</SELECT>
							<input type="Text" name="dobyear" size="4" maxlength="4" class="form" value="19" />
						</TD>
					</TR>
					<TR>
						<TD align="right" class="tdheader">
							IRC Nickname
						</TD>
						<TD class="norm">
							<input type="text" size="10" maxlength="25" name="ircnick" class="form" value="<% =Request.QueryString("ircnick") %>" />
						</TD>
					</TR>
					<TR>
						<TD align="right" class="tdheader">
							Nation
						</TD>
						<TD class="norm">
							<input type="text" size="35" name="countryname" maxlength="50" class="form" value="<% =Request.QueryString("countryname") %>" />
						</TD>
					</TR>
					<TR>
						<TD align="right" class="tdheader">
							Leader
						</TD>
						<TD class="norm">
							<input type="text" size="35" name="leadername" maxlength="50" class="form" value="<% =Request.QueryString("leadername") %>" />
						</TD>
					</TR>
					<TR>
						<TD align="right" class="tdheader" valign="top">
							Terms of Service
						</TD>
						<TD class="norm">
							<TEXTAREA cols="85" rows="12" name="toc" readonly class="form"><!---#include file="rules.txt"---></TEXTAREA>
				</table>
				<BR />
				<input type="submit" class="form" value="Register">
			</form>
		</TD>
	</tr>
</table>
<BR />
<BR />
	

<!---#include file="includes/pfooter.asp"--->
					
							