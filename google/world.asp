<html>
	<head>
		<title>Worlds of War II</title>
		<link rel=stylesheet HREF="includes/all.css">
		<meta name="description" content="Worlds of War II is a free to play, online browser, tick based strategy game.  " />
		<meta name="keywords" content="worlds, of, war, world, war, game, tick, internet, II, online, fun, free, browser, based, attack, attacking, defend, alliances, IRC, strategy, thinking, planning, multiplayer, play, players, playing" />
		<SCRIPT language="javascript">
			//<--hide script from older browsers
			//function to open a new window
			function openHelp(helpno)
			{
				//open the new window
				window.open('pop_help.asp?id='+helpno,'helppage', config='height=300,width=300,toolbar=no');
			}	
			//--end hiding>
		</script>
	</head>
<Body bgcolor="black" link="white" alink="white" vlink="white" text="#ffffff">

<center>
		<script type="text/javascript"><!--
google_ad_client = "pub-5447419161198389";
google_ad_width = 468;
google_ad_height = 60;
google_ad_format = "468x60_as";
google_ad_type = "text_image";
google_ad_channel ="";
google_page_url = document.location;
google_color_border = "336699";
google_color_bg = "FFFFFF";
google_color_link = "0000FF";
google_color_url = "008000";
google_color_text = "000000";
//--></script>
<script type="text/javascript"
  src="http://pagead2.googlesyndication.com/pagead/show_ads.js">
</script>
</center>
<BR />

<table width="95%" cellpadding="0" cellspacing="0" align="center" style="border-top: solid; border-left: solid; border-bottom: solid; border-right: solid; border-width: 1;">
	<tr>
		<td>
			<table width="100%" cellpadding="1" cellspacing="0" class="norm">
				<tr>
					<td class="tdheaderline" align="center" colspan="7">
						Lord Dave of The Flying Cheese Cake [1:1]
					</td>
				</tr>
				<tr>
					<td style="border-right: solid; border-bottom: solid; border-width: 1;" width="14%">
						Food
						<br />
						0
					</td>
					<td style="border-right: solid; border-bottom: solid; border-width: 1;" width="14%">
						Wood
						<br />
						0
					</td>
					<td style="border-right: solid; border-bottom: solid; border-width: 1;" width="15%">
						Iron
						<br />
						0
					</td>
					<td style="border-right: solid; border-bottom: solid; border-width: 1;" width="14%">
						Money
						<br />
						0
					</td>
					<td style="border-right: solid; border-bottom: solid; border-width: 1;" width="15%">
						Land
						<br />
						0
					</td>
					<td style="border-right:solid; border-bottom: solid; border-width: 1;" width="14%">
						Gold
						<br />
						0
					</td>
					<TD style="border-bottom:solid; border-width:1;" width="14%">
						Score
						<BR />
						0
					</TD>
				</tr>
				<tr>
					<td align="center" colspan="7">
						<table width="100%" cellspacing="0" cellpadding="0">
							<tr>
								<td align="left" width="30%">
									<a href="msgs.asp">
									<img src="images/mailoff.gif" border="0"></a> | <a href="priv_news.asp"><img src="images/newsoff.gif" border="0"></a>
								</td>
								</TD>
								<TD width="40%" align="center">
								</TD>
								<td width="30%" align="right" NOWRAP>
									Next Tick: 60 Minutes
								</td>
							</tr>
						</table>
					</td>
				</tr>								
			</table>
		</td>
	</tr>
</table>
<center>
<table width="95%" cellpadding="0" border="0" cellspacing="0" style="border-top:none; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD align="center" class="norm">
			<A HREF="javascript:openHelp('18');">Help With This Page</A>
		</TD>
	</TR>
</table>
</center>
<br />
<center>
	Top 10 Most Powerful Nations
</center>
<BR />
<table align="center" cellpadding="2" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<tr>
		<td class="tdheader" width="10%" style="border-right: solid; border-width:1;">
			Rank
		</td>
		<td class="tdheader" width="10%" style="border-right: solid; border-width:1;">
			Co-ordinates
		</td>
		<td class="tdheader" width="30%" style="border-right: solid; border-width:1;">
			Country
		</td>
		<td class="tdheader" width="30%" style="border-right: solid; border-width:1;">
			Leader
		</td>
		<TD class="tdheader" width="10%" style=" border-right:solid; border-width:1;">
			Land
		</TD>
		<td class="tdheader" width="10%" style="border-width:1;">
			Score
		</td>
		
	</tr>
<%
Dim adocon
Dim intCounter 		'varible used to loop through 10 records
set rsworld = server.createobject("ADODB.recordset")
set adocon = server.CreateObject("ADODB.Connection")

adocon.ConnectionString = "Provider=SQLOLEDB;Server=neptune;User ID=WoWUsr;Password=password;Database=WoWII;"

adocon.Open 

strSql = "SELECT TOP 10 tblaccounts.leadername, countryname, countryID, continent, score, amountland FROM tblAccounts ORDER BY score DESC;"

rsworld.open StrSQL, adocon

for counter = 1 to 10
	if not rsworld.eof then
%>
	<tr>
		<td class="tdheader" width="0%" style="border-top: solid; border-right: solid; border-width:1;">
			<% =counter %>
		</td>
		<td class="norm" style="border-top: solid; border-right: solid; border-width:1;">
			[<%=rsworld("continent")%>:<%=rsworld("CountryID")%>]
		</td>
		<td class="norm" style="border-top: solid; border-right: solid; border-width:1;">
			<%=rsworld("countryname")%>
		</td>
		<td class="norm" style="border-top: solid; border-right: solid; border-width:1;">
			<%=rsworld("leadername")%>
		</td>
		<TD class="norm" style="border-top:solid; border-right:solid; border-width:1;">
			<% =rsworld("amountland") %>
		</TD>
		<td class="norm" style="border-top: solid; border-width:1;">
			<%=formatnumber(rsworld("score"),0)%>
		</td>
		
	</tr>
<%
	rsworld.MoveNext
end if
next

rsworld.close
%>
</table>
<BR />
<BR />

<center>
	Top 10 Biggest Nations
</center>
<BR />
<table align="center" cellpadding="2" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<tr>
		<td class="tdheader" width="10%" style="border-right: solid; border-width:1;">
			Rank
		</td>
		<td class="tdheader" width="10%" style="border-right: solid; border-width:1;">
			Co-ordinates
		</td>
		<td class="tdheader" width="30%" style="border-right: solid; border-width:1;">
			Country
		</td>
		<td class="tdheader" width="30%" style="border-right: solid; border-width:1;">
			Leader
		</td>
		<TD class="tdheader" width="10%" style=" border-right:solid; border-width:1;">
			Land
		</TD>
		<td class="tdheader" width="10%" style="border-width:1;">
			Gold
		</td>
		
	</tr>
<%
strSql = "SELECT TOP 10 tblaccounts.leadername, countryname, countryID, continent, gold, amountland FROM tblAccounts ORDER BY amountland DESC;"

rsworld.open StrSQL, adocon

for counter = 1 to 10
	if not rsworld.eof then
%>
	<tr>
		<td class="tdheader" width="0%" style="border-top: solid; border-right: solid; border-width:1;">
			<% =counter %>
		</td>
		<td class="norm" style="border-top: solid; border-right: solid; border-width:1;">
			[<%=rsworld("continent")%>:<%=rsworld("CountryID")%>]
		</td>
		<td class="norm" style="border-top: solid; border-right: solid; border-width:1;">
			<%=rsworld("countryname")%>
		</td>
		<td class="norm" style="border-top: solid; border-right: solid; border-width:1;">
			<%=rsworld("leadername")%>
		</td>
		<TD class="norm" style="border-top:solid; border-right:solid; border-width:1;">
			<% =rsworld("amountland") %>
		</TD>
		<td class="norm" style="border-top: solid; border-width:1;">
			<%=rsworld("Gold")%>
		</td>
		
	</tr>
<%
	rsworld.MoveNext
end if
next

rsworld.close
set rsworld = nothing
%>
</table>
<BR />
<BR />
<BR />
<center>
<A HREF='http://www.ukbanners.com/cgi/clickthru.cgi?page=1' TARGET='_blank'><IMG WIDTH=468 HEIGHT=60 SRC='http://www.ukbanners.com/cgi/banners.cgi?account=B63183&page=5' BORDER=0 ALT='Click Me!'></A>
<BR />
<BR />
<font face="arial" size="2" color="#FFFFFF">
	<small>
		�<% =year(now()) %> Scott Neville, Dean Harris, James Hemming
	</small>
</div>
</center>
</body>
</html>