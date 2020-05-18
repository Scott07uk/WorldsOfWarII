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

									<img src="images/mailoff.gif" border="0">

									</a> | <a href="priv_news.asp">

									<img src="images/newsoff.gif" border="0">

									</a>
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

<!---SPELL CHECKED--->
<center>

<table width="95%" cellpadding="0" border="0" cellspacing="0" style="border-top:none; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD align="center" class="norm">
			<A HREF="javascript:openHelp('13');">Help With This Page</A>
		</TD>
	</TR>
</table>

	<BR />
	
	<table cellpadding="2" border="0" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">

		<TR>
			<TD class="tdheaderline" align="center">
				Order Defence
			</TD>
		</TR>
		<TR>
			<TD class="norm" align="center">
				Here you can order new defence to reduce the damage to your nation during attacks.
				<Br />
				
				<form action="order_defence.asp" method="post">

					<table cellpadding="2" border="0" cellspacing="0" width="100%">
						<TR>
							<TD class="tdheaderline">
								Defence Name
							</TD>
							<TD class="tdheaderline">
								Food
							</TD>
							<TD class="tdheaderline">
								Wood
							</TD>

							<TD class="tdheaderline">
								Iron
							</TD>
							<TD class="tdheaderline">
								Money
							</TD>
							<TD class="tdheaderline">
								Time
							</TD>
							<TD class="tdheaderline">

								Current Stock
							</TD>
							<TD class="tdheaderline">
								Order
							</TD>
						</TR>
						
						<TR title="Infantry unit manned in a concrete bunker.">
							<TD class="norm">
								Pillbox
							</TD>

							<TD class="norm">
								100
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								200
							</TD>
							<TD class="norm">

								500
							</Td>
							<TD class="norm">
								3 ticks
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								<input type="text" size="3" maxlength="6" name="pillboxes" class="form" value="0" />

							</TD>
						</TR>
						
						<TR title="Manned gun.">
							<TD class="norm">
								AA Gun
							</TD>
							<TD class="norm">
								100
							</TD>
							<TD class="norm">

								0
							</TD>
							<TD class="norm">
								50
							</TD>
							<TD class="norm">
								750
							</TD>
							<TD class="norm">
								4 Ticks
							</TD>

							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="aaguns" maxlength="6" class="form" />
							</TD>
						</TR>
						
						<TR title="Description.">
							<TD class="norm">

								Barrage Balloon
							</TD>
							<TD class="norm">
								100
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								50
							</TD>

							<TD class="norm">
								100
							</TD>
							<TD class="norm">
								3 ticks
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">

								<input type="text" size="3" value="0" name="balloons" maxlength="6" class="form" />
							</TD>
						</TR>
						
						<TR title="Water mines.">
							<TD class="norm">
								Water Mine
							</TD>
							<TD class="norm">
								50
							</TD>

							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								50
							</TD>
							<TD class="norm">
								100
							</TD>
							<TD class="norm">

								3 Ticks
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="mines" maxlength="6" class="form" />
							</TD>
						</TR>
						
						<TR title="Manned turret.">

							<TD class="norm">
								Turret
							</TD>
							<TD class="norm">
								100
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">

								50
							</TD>
							<TD class="norm">
								250
							</TD>
							<TD class="norm">
								4 Ticks
							</TD>
							<TD class="norm">
								0
							</TD>

							<TD class="norm">
								<input type="text" size="3" value="0" name="turrets" maxlength="6" class="form" />
							</TD>
						</TR>
						
						</table>
					<BR />
					<input type="submit" value="Order Defence" class="form">
				</form>
				
			</TD>

		</TR>
	</table>
<BR />
<BR />
	<table cellpadding="2" border="0" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
		<TR>
			<TD class="tdheaderline" align="center">
				Incoming Units
			</TD>
		</TR>

		<TR>
			<TD class="norm" align="center">
				Below is a table of all the units you have ordered.
				<BR />
				<table cellpadding="2" cellspacing="0" border="0">
					<TR>
						<TD class="tdheaderline">
							Unit Name
						</TD>
						<TD class="tdheaderline" align="center" width="17px">1</TD><TD class="tdheaderline" align="center" width="17px">2</TD><TD class="tdheaderline" align="center" width="17px">3</TD><TD class="tdheaderline" align="center" width="17px">4</TD>

						
					</TR>
					
					<TR>
					<TR><TD class=form align=center>Pillboxes</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR><TR><TD class=form align=center>AA Guns</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR><TR><TD class=form align=center>Barrage Balloons</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR><TR><TD class=form align=center>Mines</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR><TR><TD class=form align=center>Turrets</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR>
					</TR>
				</table>
			</TD>

		</TR>
	</table>
	<BR />
</center>
<BR />
<center>

<A HREF='http://www.ukbanners.com/cgi/clickthru.cgi?page=1' TARGET='_blank'><IMG WIDTH=468 HEIGHT=60 SRC='http://www.ukbanners.com/cgi/banners.cgi?account=B63183&page=5' BORDER=0 ALT='Click Me!'></A>
<BR />
<BR />

<font face="arial" size="2" color="#FFFFFF">
	<small>

		©2005 Scott Neville, Dean Harris, James Hemming
	</small>
</div>
</center>
</body>
</html>