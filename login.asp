<!---SPELL CHECKED--->

<HTML>
	<HEAD>
		<TITLE>Worlds of War II Login</TITLE>
		<link href="styles/default.css" rel="stylesheet" type="text/css">
	</HEAD>
	
	<BODY bgcolor="#000000">
	<center>
		<table cellpadding="2" cellspacing="0" border="0">
			<TR>
				<TD align="center" style="border-bottom:solid; border-color:#FFFFFF; border-width:1;">
					<img SRC="images/logo.gif" border="0" alt="Worlds of War II" />
				</TD>
			</TR>
			<TR>
				<TD align="center" style="border-left:solid; border-right:solid; border-bottom:solid; border-color:#FFFFFF; border-width:1;">
					<font face="arial" size="2" color="#FFFFFF">
						<B>
							Login to Worlds of War II
						</B>
					</font>
					<BR />
					<font face="arial" size="2" color="#FFFFFF">
					<%
					select case Request.QueryString("error")
						case 1
							Response.Write("The username or password you entered is invalid.<BR /><BR />")
						case 2
							Response.Write("This account has been suspended. Please check your email to find out why.<BR /><BR />")
						case 4
							Response.Write("You did not enter an email address.<BR /><BR />")
						case 5
							Response.Write("The email address you supplied is not in our database.<BR /><BR />")
						case 6
							Response.Write("Your login details have been emailed to you. Please note the email may be filtered as spam.<BR /><BR />")
					end select
					%>
					</font>
					<form action="login_engine.asp" method="post">
						<table cellpadding="2" cellspacing="0" border="0">
							<TR>
								<TD align="right">
									<font face="arial" size="2" color="#FFFFFF">
										Username
									</font>
								</TD>
								<TD>
									<input type="text" size="15" name="username" maxlength="50" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-width:1; border-color:#000000;" />
								</TD>
							</TR>
							<TR>
								<TD align="right">
									<font face="arial" size="2" color="#FFFFFF">
										Password
									</font>
								</TD>
								<TD>
									<input type="password" size="15" name="password" maxlength="50" style="border-top:solid; border-left:solid; border-right:solid; border-bottom:solid; border-width:1; border-color:#000000;" />
								</TD>
							</TR>
						</table>
						<BR>
						<input type="submit" value="Login" style="border-top:solid; border-bottom:solid; border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;" />
					</form>
				</TD>
			</TR>
			<TR>
				<TD align="center" style="border-left:solid; border-right:solid; border-bottom:solid; border-width:1; border-color:#FFFFFF;">
					<font face="arial" size="2" color="#FFFFFF">
						<B>
							Recover Password
						</B>
						<BR />
						<BR />
						To recover a lost password, enter your registered email address in the box below and it will be emailed to you.
						<BR/>
						<form action="recover_password_engine.asp" method="post">
							<input type="text" size="30" name="email" maxlength="150" style="border-top:solid; border-left:solid; border-right:solid; border-bottom:solid; border-color:#000000; border-width:1;" />
							<BR />
							<BR />
							<input type="submit" value="Recover Password" style="border-top:solid; border-left:solid; border-right:solid; border-bottom:solid; border-color:#FFFFFF; border-width:1;" />
						</form>
					</font>
				</TD>
			</TR>
		</table>
	</center>
	</body>
</HTML>
					
			