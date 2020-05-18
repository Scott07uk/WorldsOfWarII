<!---#include file="includes/ptop.asp"--->

<!---SPELL CHECKED--->

<%
select case Request.QueryString("type")
	case "cheet"
%>
<center>
	<table cellpadding="2" cellspacing="0" border="0" width="95%" style="border-left:solid; border-right:solid; border-top:solid; border-bottom:solid; border-width:1; border-color:#FFFFFF;">
	
		<TR>
			<TD class="tdheader" align="center">
				Report Cheating
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				If you suspect someone of cheeting in the game, please use the form below to report them to the admins so that the necessary action can be taken.
				<BR />
				<BR />
				<center>
					<form action="report_cheeting_engine.asp" method="post">
						<table cellpadding="2" border="0" cellspacing="0" style="border-top:solid; border-left:solid; border-right:solid; border-bottom:solid; border-color:#FFFFFF; border-width:1;">
							<TR>
								<TD class="tdheader" align="right" valign="top" width="175">
									Type of Offence
								</TD>
								<TD class="norm">
									<select name="type" size="1" class="form">
										<option value="Multiple Accounts">Multiple Accounts</option>
										<option value="Game Bug">Game Bug</option>
										<option value="Hacking">Hacking</option>
										<option value="Other">Other</option>
									</select>
								</TD>
							</TR>
							<TR>
								<TD class="tdheader" align="right" valign="top">
									Your Username
								</TD>
								<TD class="norm">
									<input type="text" size="20" name="username" maxlength="50" class="form" />
								</TD>
							</TR>
							<TR>
								<TD class="tdheader" align="right" valign="top">
									Location of Cheating Account
								</TD>
								<TD class="norm">
									<input type="text" name="continent" size="3" maxlength="4" class="form" />
									:
									<input type="text" name="country" size="2" maxlength="2" class="form" />
								</TD>
							</TR>
							<TR>
								<TD class="tdheader" align="right" valign="top" width="175">
									Description of Offence
								</TD>
								<TD class="norm">
									<TEXTAREA cols="50" rows="5" name="description" class="form"></textarea>
								</TD>
							</TR>
							<TR>
								<TD class="norm">&nbsp;
								
								</TD>
								<TD class="norm">
									<input type="submit" value="Post Report" class="form" />
								</TD>
							</TR>
						</Table>
					</form>
				</center>
			</TD>
		</TR>
	</table>
</center>
	
<%
	case "abuse"
	%>
<center>
	<table cellpadding="2" cellspacing="0" border="0" width="95%" style="border-left:solid; border-right:solid; border-top:solid; border-bottom:solid; border-width:1; border-color:#FFFFFF;">
	
		<TR>
			<TD class="tdheader" align="center">
				Report Game Abuse
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				If you find someone being abusive in the game, please report it here so we can take the necessary action.
				<BR />
				<BR />
				<center>
					<form action="report_abuse_engine.asp" method="post">
						<table cellpadding="2" border="0" cellspacing="0" style="border-top:solid; border-left:solid; border-right:solid; border-bottom:solid; border-color:#FFFFFF; border-width:1;">
							<TR>
								<TD class="tdheader" align="right" valign="top" width="175">
									Type of Abuse
								</TD>
								<TD class="norm">
									<select name="type" size="1" class="form">
										<option value="Spam">Spam</option>
										<option value="Message abuse">Abuse in Messages</option>
										<option value="Forum abuse">Abuse on the Forum</option>
										<option value="Other">Other</option>
									</select>
								</TD>
							</TR>
							<TR>
								<TD class="tdheader" align="right" valign="top">
									Your Username
								</TD>
								<TD class="norm">
									<input type="text" size="20" name="username" maxlength="50" class="form" />
								</TD>
							</TR>
							<TR>
								<TD class="tdheader" align="right" valign="top">
									Location of Offending Account
								</TD>
								<TD class="norm">
									<input type="text" name="continent" size="3" maxlength="4" class="form" />
									:
									<input type="text" name="country" size="2" maxlength="2" class="form" />
								</TD>
							</TR>
							<TR>
								<TD class="tdheader" align="right" valign="top" width="175">
									Description of Offence (Including Subjects of Offending Messages)
								</TD>
								<TD class="norm">
									<TEXTAREA cols="50" rows="5" name="description" class="form"></textarea>
								</TD>
							</TR>
							<TR>
								<TD class="norm">&nbsp;
								
								</TD>
								<TD class="norm">
									<input type="submit" value="Post Report" class="form" />
								</TD>
							</TR>
						</Table>
					</form>
				</center>
			</TD>
		</TR>
	</table>
</center>
	

<%

case "suggestion"
%>
<center>
	<table cellpadding="2" cellspacing="0" border="0" width="95%" style="border-left:solid; border-right:solid; border-top:solid; border-bottom:solid; border-width:1; border-color:#FFFFFF;">
	
		<TR>
			<TD class="tdheader" align="center">
				Make Suggestion
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				If you have an idea to help improve Worlds of War II, use this form to let us know.
				<BR />
				<BR />
				<center>
					<form action="report_idea_engine.asp" method="post">
						<table cellpadding="2" border="0" cellspacing="0" style="border-top:solid; border-left:solid; border-right:solid; border-bottom:solid; border-color:#FFFFFF; border-width:1;">
							<TR>
								<TD class="tdheader" align="right" valign="top" width="175">
									Applies To
								</TD>
								<TD class="norm">
									<select name="type" size="1" class="form">
										<option value="Portal">Portal</option>
										<option value="Announcements">Announcements</option>
										<option value="Registration">Registration</option>
										<option value="Game Princibles">Game Principles</option>
										<option value="Messages">Messages</option>
										<option value="Continent Forum">Continent Forum</option>
										<option value="News">News</option>
										<option value="IRC">IRC</option>
										<option value="Research">Research</option>
										<option value="Construction">Construction</option>
										<option value="Units">Units</option>
										<option value="Politics">Politics</option>
										<option value="Ranking">Ranking</option>
										<option value="Attacking">Attacking</option>
										<option value="Defending">Defending</option>
										<option value="Forums">Forums</option>
										<option value="Other">Other</option>
									</select>
								</TD>
							</TR>
							<TR>
								<TD class="tdheader" align="right" valign="top">
									Your Email
								</TD>
								<TD class="norm">
									<input type="text" size="40" name="email" maxlength="150" class="form" />
								</TD>
							</TR>
							<TR>
								<TD class="tdheader" align="right" valign="top">
									Idea Title
								</TD>
								<TD class="norm">
									<input type="text" name="title" size="30" maxlength="50" class="form" />
								</TD>
							</TR>
							<TR>
								<TD class="tdheader" align="right" valign="top" width="175">
									Detailed Description of Idea
								</TD>
								<TD class="norm">
									<TEXTAREA cols="65" rows="7" name="description" class="form"></textarea>
								</TD>
							</TR>
							<TR>
								<TD class="norm">&nbsp;
								
								</TD>
								<TD class="norm">
									<input type="submit" value="Post Report" class="form" />
								</TD>
							</TR>
						</Table>
					</form>
				</center>
			</TD>
		</TR>
	</table>
</center>
<%
end select
%>
<!---#include file="includes/pfooter.asp"--->