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
			<form action="add_email_engine.asp" method="post">
				<input type="hidden" name="redir" value="add_email2.asp">
				<%
				select case Request.QueryString("error")
					case 1
						Response.Write("You have not entered an email address to activate. <BR /><BR />")
					case 2
						Response.Write("The email address you entered is not valid, please enter a valid email address to continue. <BR /><BR />")
					case 3
						Response.Write("Sorry but this email address is banned from use on this site.  <BR /><BR />")
					case 4
						Response.Write("Sorry but email addresses using that domain are banned from use on this site.  <BR /><BR />")
					case 5
						Response.Write("Sorry but the email address you entered is already in our database, if you want to activate this email use the link below.   <BR /><BR />")
				end select
				%>
				In order to use an email address with Worlds of War II you need get it activated. 
				<BR />
				Please enter your email address below. You will be emailed a code which you will need to enter on the next page to confirm that the address is valid.  
				<BR />
				<BR />
				If you already have an activation code for an email address, click <A HREF="add_email2.asp">here</A>.  
				<BR />
				<input type="text" size="30" maxlength="100" name="email" class="form">
				<BR />
				<BR />
				<input type="submit" value="Submit Address" class="form">
			</form>
		</TD>
	</TR>
</table>
</center>
</body>
</html>