<!---#include file="includes/ptop.asp"--->

<!---SPELL CHECKED--->

<table cellpadding="5" border="0" cellspacing="5" style="border-left:solid; border-right:solid; border-top:solid; border-bottom:solid; border-width:1; border-color:#FFFFFF;" width="95%">
	<TR>
		<TD class="tdheader" align="center">
			Current Status:
		</TD>
	</TR>
	<%
	Dim rsserver
	Dim rsstat
	Dim lngNumAccounts
	Dim lngTotalLand
	Dim lngTotalGold
	Dim strLargePlayer
	Dim lngNumContinents
	Dim lngInvasions
	
	lngNumAccounts = 0
	lngTotalLand = 0
	lngtotalGold = 0
	lngNumContinents = 0
	lngInvasions = 0
		
	Set adoCon = Server.CreateObject("ADODB.Connection")
		
	adocon.ConnectionString = "Provider=SQLOLEDB;Server=neptune;User ID=WoWUsr;Password=password;Database=WoWII;"
		
	adocon.Open
	
	set rsstat = server.CreateObject("ADODB.Recordset")
	
	strsql = "SELECT amountland, gold, leadername, countryname FROM tblaccounts ORDER by amountland DESC;"
	
	rsstat.Open strsql, adocon
	
	if rsstat.EOF = false then
			strlargePlayer = rsstat("leadername") & " of " & rsstat("countryname")
	end if
	
	do while not rsstat.EOF
		lngNumAccounts = lngNumAccounts + 1
		lngTotalLand = clng(lngtotalland) + clng(rsstat("amountland"))
		lngTotalGold = clng(lngTotalGold) + clng(rsstat("gold"))
	
		rsstat.MoveNext 
	loop
	
	rsstat.Close
	
	strsql = "SELECT continentno FROM tblcontinents;"
	
	rsstat.Open strsql, adocon
	
	do while not rsstat.EOF 
		lngNumContinents = lngNumContinents + 1
		rsstat.Movenext
	loop
	
	rsstat.Close 
	
	strsql = "SELECT status FROM tblmovements WHERE status = 1;"
	
	rsstat.Open strsql, adocon
	do while not rsstat.EOF
		lnginvasions = lnginvasions + 1
		
		rsstat.MoveNext
	loop
	
	rsstat.Close 
	
	set rsstat = nothing
	%>
	
	<TR>
		<TD class="norm">
			<B>Worlds of War II Status</B>
			<table cellpadding="2" border="0" cellspacing="0">
				<TR>
					<TD class="norm" align="right">
						Number of Accounts
					</TD>
					<TD class="norm">
						<% =lngNumAccounts %>
					</TD>
				</TR>
				<TR>
					<TD class="norm" align="right">
						Number of Continents
					</TD>
					<TD class="norm">
						<% =lngNumContinents %>
					</TD>
				</TR>
				<TR>
					<TD class="norm" align="right">
						Total Land Discovered
					</TD>
					<TD class="norm">
						<% =formatnumber(lngtotalland,0) %>
					</TD>
				</TR>
				<TR>	
					<TD align="right" class="norm">
						Total Gold Discovered
					</TD>
					<TD class="norm">
						<% =lngtotalgold %>
					</TD>
				</TR>
				<TR>
					<TD align="right" class="norm">
						Largest Player
					</TD>
					<TD class="norm">
						<% =strlargeplayer %>
					</TD>
				</TR>
				<TR>
					<TD align="right" class="norm">
						Number of Invasions in Progress
					</TD>
					<TD class="norm">
						<% =lnginvasions %>
					</TD>
				</TR>
			</table>
			<BR />
			<BR />
			<B>Network Status</B>
			<%
			
			strSql = "SELECT address, status, networkbots, statsbot, allowconnect FROM tblircservers WHERE status <> 3;"
			
			set rsserver = server.CreateObject("ADODB.Recordset")
			
			rsserver.Open strsql, adocon
			%>
			<table cellpadding="2" border="0" cellspacing="0">
			<%
			do while not rsserver.EOF 
			%>
				<TR>
					<TD class="norm">
					<%
						if cbool(rsserver("allowconnect")) = true then
					%>
						<A HREF="irc://<% =rsserver("address") %>:6667"><% =rsserver("address") %></A>
					<%
					else
					if cbool(rsserver("networkbots")) = true then
					%>
					<% =rsserver("address") %> (Network Services)
					<%
					else
					if cbool(rsserver("statsbot")) = true then
					%>
					<% =rsserver("address") %> (Statistics Services)
					<%
					else
					%>
					<% =rsserver("address") %>
					<%
					end if
					end if
					end if
					%>
					</TD>
					<TD class="norm">
					<%
					if rsserver("status") = 1 then
					%>
						<img src="images/tick.gif" border="0" alt="All Go" width="25" height="25" />
					<%
					else
					%>
						<img src="images/cross.gif" border="0" alt="Offline" width="25" height="25" />
					<%
					end if
					%>
					</TD>
				</TR>
				<%
				if rsserver("address") = "irc.worldsofwar.net" then
						mainircstatus = rsserver("status")
				end if
				rsserver.MoveNext 
				loop
				
				strsql = "SELECT ticksenabled FROM tblconfig;"
				
				rsserver.Close 
				
				rsserver.Open strsql, adocon
				
				%>
				<TR>
					<TD class="norm">
						saturn.worldsofwar.co.uk (Tick Engine)
					</TD>
					<TD class="norm">
						<% if cbool(rsserver("ticksenabled")) = true then %>
						<img src="images/tick.gif" border="0" alt="All Go" width="25" height="25" />
						<% else %>
						<img src="images/cross.gif" border="0" alt="Offline" width="25" height="25" />
						<% end if %>
					</TD>
				</TR>
				<TR>
					<TD class="norm">
						<A HREF="http://www.worldsofwar.co.uk" target="_blank">neptune.worldsofwar.co.uk (HTTP Server)</A>
					</TD>
					<TD class="norm">
						<img src="images/tick.gif" border="0" alt="All Go" width="25" height="25" />
					</TD>
				</TR>
				<TR>
					<TD class="norm">
						wowstats.worldsofwar.net (IRC Information Bot)
					</TD>
					<TD class="norm">
						<img src="images/cross.gif" border="0" alt="Offline" width="25" height="25" />
					</TD>
				</TR>
				<TR>
					<TD class="norm">
						<A HREF="https://www.worldsofwar.net" target="_blank">pluto.worldsofwar.net (IRC Services Website)</A>
					</TD>
					<TD class="norm">
					<%
					if mainircstatus = 1 then
					%>
						<img src="images/tick.gif" border="0" alt="All Go" width="25" height="25" />
					<%
					else
					%>
						<img src="images/cross.gif" border="0" alt="Offline" width="25" height="25" />
					<% end if %>
					</TD>
				</TR>
			</table>
			<%
			rsserver.close
			adocon.Close
			set rsserver = nothing
			set adocon = nothing
			%>
			<BR />
			<Br />
			<B>Overall Game Ranking</B>
			<BR />
			<SCRIPT language="Javascript" src="http://www.topwebgames.com/games/activead.js?id=3239"></SCRIPT>
			<BR />
			<BR />
			<center>
			<A href="http://www.topwebgames.com/in.asp?id=3239" target="_blank">Support Worlds of War II. Click here to vote for Worlds of War II each week.</A>
			</center>
		</TD>
	</TR>
</table>
<BR />
<BR />
	

<!---#include file="includes/pfooter.asp"--->