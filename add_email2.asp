<!---SPELL CHECKED--->

<HTML>
	<HEAD>
		<TITLE>
			--==Welcome to Worlds of War II==--
		</TITLE>
		<link rel=stylesheet HREF="includes/all.css">
	</HEAD>
	<Body bgcolor="black" link="white" alink="white" vlink="white" text="#ffffff" topmargin="0" bottommargin="0" leftmargin="0" rightmargin="0">
	<center>
	<table cellpadding="5" width="95%" border="0" cellspacing="5" style="border-left:solid; border-right:solid; border-top:solid; border-bottom:solid; border-width:1; border-color:#FFFFFF;" width="95%">
	<TR>
		<TD class="tdheader" align="center">
			Activate Email
		</TD>
	</TR>
	<TR>
		<TD class="norm" align="center">
			<form action="activate_email_engine.asp" method="post">
				<%
				select case Request.QueryString("error")
					case 1
						Response.Write("You did not enter an email address to activate.   <BR /><BR />")
					case 2
						Response.Write("The email address you entered is not stored in our database. <BR /><BR />")
					case 3
						Response.Write("The code you entered does not match the email address stored.  <BR /><BR />")
				end select
				%>
				Once you have received your activation email you can fill in the form below.  
				<BR />
				<BR />
				The email may be filtered as junk mail if you use a free email service such as Yahoo or Hotmail.  
				<BR />
				<BR />
				If you need another copy of the email, please enter your email address in the space below and leave the activation code blank.  
				<BR />
				<BR />
				<table cellpadding="2" cellspacing="0" border="0">
					<TR>
						<TD align="right" class="norm">
							Email address:
						</TD>
						<TD class="norm">
							<input type="text" size="30" name="email" class="form" maxlength="150">
						</TD>
					</TR>
					<TR>
						<TD align="right" class="norm">
							Activation code:
						</TD>
						<TD class="norm">
							<input type="text" size="30" name="code" class="form" maxlength="50">
						</TD>
					</TR>
				</table>
				<BR />
				<input type="submit" value="Activate Email" class="form">
			</form>
		</TD>
	</TR>
</table>
</center>
</body>
</html>