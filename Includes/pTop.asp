<HTML>
	<HEAD>
		<TITLE>
			Worlds of War II
		</TITLE>
		<meta name="description" content="Worlds of War II is a free to play, online browser, tick based strategy game.  " />
		<meta name="keywords" content="worlds, of, war, world, war, game, tick, internet, II, online, fun, free, browser, based, attack, attacking, defend, alliances, IRC, strategy, thinking, planning, multiplayer, play, players, playing" />
		<meta name="rating" content="general" />
		<meta name="copyright" content="Copyright �2005 - Scott Neville" />
		<meta name="revisit-after" content="7 Days" />
		<meta name="expires" content="never">
		<meta name="distribution" content="global" />
		<meta name="robots" content="index,follow" />
		<meta name="author" content="Scott Neville, Dean Harris & James Hemming" />
		<link rel=stylesheet HREF="includes/all.css">
		<SCRIPT language="javascript">
		//<--
		
		function processUnix(){
		var strEntered
		
		strEntered = document.unix.input.value;
		
		strEntered = strEntered.toLowerCase();
		
		switch (document.unix.input.value)
		{
		case "home":location.href='default.asp';break;
		case "news":location.href='announcements.asp';break;
		case "search":location.href='search.asp';break;
		case "status":location.href='status.asp';break;
		case "register":location.href='register1.asp';break;
		case "login":location.href='login.asp';break;
		case "help":location.href='help.asp';break;
		case "cheeting":location.href='report_from.asp?type=cheet';break;
		case "abuse":location.href='report_form.asp?type=abuse';break;
		case "ideas":location.href='report_form.asp?type=suggestion';break;
		case "manual":window.open('manual/default.asp','manual');break;
		case "irc":window.open('https://worldsofwar.co.uk','IRC');break;
		case "forums":window.open('http://www.fhforums.co.uk','Forums');break;
		case "credits":location.href='credits.asp';break;
		case "halt":window.close();break;
		
		}
		
		document.unix.input.value="";
		return false;
		}
		
		
		
		//-->
		</script>
	</HEAD>
	<%
	Dim adocon
	Dim rsconfig
	
	set adocon1 = server.CreateObject("ADODB.Connection")
	
	adocon1.ConnectionString = "Provider=SQLOLEDB;Server=neptune;User ID=WoWUsr;Password=password;Database=WoWII;"
	
	adocon1.Open 
	
	strsql = "SELECT portaladdress, forumsaddress, manualaddress, ircsite FROM tblconfig;"
	
	set rsconfig = server.CreateObject("ADODB.Recordset")
	
	rsconfig.Open strsql, adocon1
	%>
	<Body bgcolor="black" link="white" alink="white" vlink="white" text="#ffffff" topmargin="0" bottommargin="0" leftmargin="0" rightmargin="0">
	<table cellpadding="0" border="0" cellspacing="0" width="100%">
		<TR>
			<TD align="center" class="norm">
				
					<form action="login_engine.asp" method="post">
						Quick Login:
						&nbsp;
						&nbsp;
						Username:<input type="text" name="username" size="15" maxlength="50" class="form">
						&nbsp;
						Password:<input type="password" name="password" size="15" maxlength="50" class="form">
						&nbsp;
						<input type="submit" value="Login" class="form">
						&nbsp;
						&nbsp;
						<A HREF="login.asp" target="_top">Forgotton your password?</A>
					</form>
				
			</TD>
		</TR>
		<TR>
			<TD align="center">
				<table cellpadding="0" border="0" cellspacing="0" width="100%">
					<TR>
						<TD align="left" width="150px" valign="top" class="plinks">
							<small>
								WOWII su:game to root on /dev/ttyp0
							</small>
							<BR />
							<BR />
							<A HREF="default.asp" name="linkhome" class="plinks">WOWII# Home</A>
							<BR />
							<A HREF="announcements.asp" class="plinks">WOWII# News</A>
							<BR />
							<A HREF="search.asp" class="plinks">WOWII# Search</A>
							<BR />
							<A HREF="Status.asp" class="plinks">WOWII# Status</A>
							<BR />
							<BR />
							<A HREF="register1.asp" class="plinks">WOWII# Register</A>
							<BR />
							<A HREF="login.asp" class="plinks">WOWII# Login</A>
							<BR />
							<A HREF="Help.asp" class="plinks">WOWII# Help</A>
							<BR />
							<BR />
							<A HREF="report_form.asp?type=cheet" class="plinks">WOWII# Cheating</A>
							<BR />
							<A HREF="report_form.asp?type=abuse" class="plinks">WOWII# Abuse</A>
							<BR />
							<A HREF="report_form.asp?type=suggestion" class="plinks">WOWII# Ideas</A>
							<BR />
							<BR />
							<A HREF="<% =rsconfig("manualaddress") %>" target="_blank" class="plinks">WOWII# Manual</A>
							<BR />
							<A HREF="<% =rsconfig("ircsite") %>" target="_blank" class="plinks">WOWII# IRC</A>
							<BR />
							<A HREF="<% =rsconfig("forumsaddress") %>" target="_blank" class="plinks">WOWII# Forums</A>
							<BR />
							<BR />
							<A HREF="credits.asp" class="plinks">WOWII# Credits</a>
							<br />
							<br />
							<form name="unix" onsubmit="return processUnix();" >
								<input type="text" size="15" maxlength="20" name="input" class="unixform" />
							</form>
							<span name="bob">
							</span>
						</TD>
						<TD valign="top" align="center">
						<%
						rsconfig.close
						strsql = "SELECT adText, adshows FROM tblBanners WHERE Active = 1 ORDER by adshows ASC;"
										
						rsconfig.LockType = 3 
						rsconfig.CursorType = 2
										
						rsconfig.open strsql, adocon1
						if not rsconfig.EOF then
								call rsconfig.Update("adshows", clng(rsconfig("adshows"))+1)
								%><BR />
								<%
								'Response.write(rsconfig("adtext")) 
								%><BR />
								
						<%
						end if
						%>
							<BR />
							<%
							rsconfig.close
							adocon1.Close 
							
							set rsconfig = nothing
							set adocon1 = nothing
							%>