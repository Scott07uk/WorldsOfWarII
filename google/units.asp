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
			<A HREF="javascript:openHelp('12');">Help With This Page</A>
		</TD>
	</TR>
</table>
</center>
<center>
	<BR />

	
	<table cellpadding="2" border="0" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
		<TR>
			<TD class="tdheaderline" align="center">
				Order Units
			</TD>
		</TR>
		<TR>
			<TD class="norm" align="center">
				Here you can order new units to attack and defend yourself.
				<Br />

				
				<form action="order_units.asp" method="post">
					<table cellpadding="2" border="0" cellspacing="0" width="100%">
						<TR>
							<TD class="tdheaderline">
								Unit Name
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
						<TR title="Infantry are the smallest and most common unit on the battlefield.">
							<TD class="norm">
								Infantry
							</TD>

							<TD class="norm">
								50
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								5
							</TD>
							<TD class="norm">

								100
							</Td>
							<TD class="norm">
								3 Ticks
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								<input type="text" size="3" maxlength="6" name="infantry" class="form" value="0" />

							</TD>
						</TR>
						
						<TR title="Commandos are another form of ground troop that are better trained and armed.">
							<TD class="norm">
								Commandos
							</TD>
							<TD class="norm">
								75
							</TD>
							<TD class="norm">

								0
							</TD>
							<TD class="norm">
								10
							</TD>
							<TD class="norm">
								150
							</TD>
							<TD class="norm">
								4 Ticks
							</TD>

							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="comandoes" maxlength="6" class="form" />
							</TD>
						</TR>
						
						<TR title="The Jeep is the cheapest motorised unit. It is highly mobile and has mounted guns.">
							<TD class="norm">

								Jeep
							</TD>
							<TD class="norm">
								100
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								75
							</TD>

							<TD class="norm">
								300
							</TD>
							<TD class="norm">
								6 Ticks
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">

								<input type="text" size="3" value="0" name="jeeps" maxlength="6" class="form" />
							</TD>
						</TR>
						
						<TR title="The Matilda MkI is basically a slow moving excessively well armoured mobile machine gun.">
							<TD class="norm">
								Matilda MkI
							</TD>
							<TD class="norm">
								150
							</TD>

							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								100
							</TD>
							<TD class="norm">
								500
							</TD>
							<TD class="norm">

								8 Ticks
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="matilda1" maxlength="6" class="form" />
							</TD>
						</TR>
						
						<TR title="The Matilda MkII is a mobile fortress with a sound 2lb turret mounted gun.">

							<TD class="norm">
								Matilda MkII
							</TD>
							<TD class="norm">
								150
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">

								120
							</TD>
							<TD class="norm">
								600
							</TD>
							<TD class="norm">
								8 Ticks
							</TD>
							<TD class="norm">
								0
							</TD>

							<TD class="norm">
								<input type="text" size="3" value="0" name="matilda2" maxlength="6" class="form" />
							</TD>
						</TR>
						
						<TR title="The Sherman Tank is a faster, more powerful tank than the Matilda's but lacks armour.">
							<TD class="norm">
								Sherman Tank
							</TD>
							<TD class="norm">

								150
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								150
							</TD>
							<TD class="norm">
								700
							</TD>

							<TD class="norm">
								10 Ticks
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="sherman" maxlength="4" class="form" />
							</TD>
						</TR>

						
						<TR title="The Spitfire MkI is the first high speed, all metal single wing fighter plane.">
							<TD class="norm">
								Supermarine Spitfire MkI
							</TD>
							<TD class="norm">
								200
							</TD>
							<TD class="norm">
								10
							</TD>

							<TD class="norm">
								100
							</TD>
							<TD class="norm">
								700
							</TD>
							<TD class="norm">
								8 Ticks
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="spitfire1" maxlength="4" class="form" />
							</TD>
						</TR>
						
						<TR title="The MkII Spitfire now features air to air cannons with improved flight performance.">
							<TD class="norm">
								Supermarine Spitfire MkII
							</TD>

							<TD class="norm">
								200
							</TD>
							<TD class="norm">
								10
							</TD>
							<TD class="norm">
								125
							</TD>
							<TD class="norm">

								750
							</TD>
							<TD class="norm">
								10 Ticks
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="spitfire2" maxlength="4" class="form" />

							</TD>
						</TR>
						
						<TR title="The MkV Spitfire now incorporates under wing rockets for strafing ground targets.">
							<TD class="norm">
								Supermarine Spitfire MkV
							</TD>
							<TD class="norm">
								250
							</TD>
							<TD class="norm">

								10
							</TD>
							<TD class="norm">
								150
							</TD>
							<TD class="norm">
								800
							</TD>
							<TD class="norm">
								12 Ticks
							</TD>

							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="spitfire5" maxlength="4" class="form" />
							</TD>
						</TR>
						
						<TR title="The Hawker Hurricane is a new type of single wing plane. It is very durable and quick to repair.">
							<TD class="norm">

								Hawker Hurricane MkI
							</TD>
							<TD class="norm">
								200
							</TD>
							<TD class="norm">
								50
							</TD>
							<TD class="norm">
								75
							</TD>

							<TD class="norm">
								600
							</TD>
							<TD class="norm">
								7 Ticks
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">

								<input type="text" size="3" value="0" name="hurricane" maxlength="4" class="form" />
							</TD>
						</TR>
						
						<TR title="The Halifax Bomber is a slow cumbersome long-range bomber.">
							<TD class="norm">
								Halifax Bomber
							</TD>
							<TD class="norm">
								300
							</TD>

							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								200
							</TD>
							<TD class="norm">
								900
							</TD>
							<TD class="norm">

								10 Ticks
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="halifax" maxlength="4" class="form" />
							</TD>
						</TR>
						
						<TR title="The Wellington Bomber is a more modern, adaptable medium range bomber.">

							<TD class="norm">
								Wellington Bomber
							</TD>
							<TD class="norm">
								300
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">

								200
							</TD>
							<TD class="norm">
								1000
							</TD>
							<TD class="norm">
								12 Ticks
							</TD>
							<TD class="norm">
								0
							</TD>

							<TD class="norm">
								<input type="text" size="3" value="0" name="wellington" maxlength="4" class="form" />
							</TD>
						</TR>
						
						<TR title="The Lancaster is a tough, long range bomber that uses the G navigation system for extra accuracy.">
							<TD class="norm">
								Lancaster Bomber
							</TD>
							<TD class="norm">

								350
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								300
							</TD>
							<TD class="norm">
								1100
							</TD>

							<TD class="norm">
								15 Ticks
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="lancaster" maxlength="4" class="form" />
							</TD>
						</TR>

						
						<TR title="The Landing Craft is used to transport troops to enemy shores for invasion.">
							<TD class="norm">
								Landing Craft
							</TD>
							<TD class="norm">
								75
							</TD>
							<TD class="norm">
								0
							</TD>

							<TD class="norm">
								90
							</TD>
							<TD class="norm">
								150
							</TD>
							<TD class="norm">
								5 Ticks
							</TD>
							<TD class="norm">

								0
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="landingcraft" maxlength="4" class="form" />
							</TD>
						</TR>
						
						<TR title="The Landing Ship can be used to take tanks and other motorised units to enemy shores.">
							<TD class="norm">
								Landing Ship
							</TD>

							<TD class="norm">
								150
							</TD>
							<TD class="norm">
								10
							</TD>
							<TD class="norm">
								500
							</TD>
							<TD class="norm">

								300
							</TD>
							<TD class="norm">
								10 Ticks
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="landingship" maxlength="4" class="form" />

							</TD>
						</TR>
						
						<TR title="The Frigates are smaller war ships often used as escorts to assist the main ships.">
							<TD class="norm">
								Frigate
							</TD>
							<TD class="norm">
								250
							</TD>
							<TD class="norm">

								20
							</TD>
							<TD class="norm">
								1000
							</TD>
							<TD class="norm">
								1100
							</TD>
							<TD class="norm">
								16 Ticks
							</TD>

							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="frigates" maxlength="4" class="form" />
							</TD>
						</TR>
						
						<TR title="The pride of the navy, the large Battleships, are designed for sinking its own class and decimating coastlines.">
							<TD class="norm">

								Battleship
							</TD>
							<TD class="norm">
								400
							</TD>
							<TD class="norm">
								300
							</TD>
							<TD class="norm">
								2000
							</TD>

							<TD class="norm">
								1500
							</TD>
							<TD class="norm">
								20 Ticks
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">

								<input type="text" size="3" value="0" name="battleship" maxlength="4" class="form" />
							</TD>
						</TR>
						
						<TR title="The navy's largest ships, used to reshape coastlines and handle combat with several Frigates at a time.">
							<TD class="norm">
								Aircraft Carrier
							</TD>
							<TD class="norm">
								500
							</TD>

							<TD class="norm">
								400
							</TD>
							<TD class="norm">
								3000
							</TD>
							<TD class="norm">
								2000
							</TD>
							<TD class="norm">

								24 Ticks
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="carriers" maxlength="4" class="form" />
							</TD>
						</TR>
						
					</table>

					<BR />
					<input type="submit" value="Order Units" class="form">
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
				Below is a table detailing all the units you have ordered.
				<BR />
				<table cellpadding="2" cellspacing="0" border="0">
					<TR>

						<TD class="tdheaderline">
							Unit Name
						</TD>
						<TD class="tdheaderline" align="center" width="17px">1</TD><TD class="tdheaderline" align="center" width="17px">2</TD><TD class="tdheaderline" align="center" width="17px">3</TD><TD class="tdheaderline" align="center" width="17px">4</TD><TD class="tdheaderline" align="center" width="17px">5</TD><TD class="tdheaderline" align="center" width="17px">6</TD><TD class="tdheaderline" align="center" width="17px">7</TD><TD class="tdheaderline" align="center" width="17px">8</TD><TD class="tdheaderline" align="center" width="17px">9</TD><TD class="tdheaderline" align="center" width="17px">10</TD><TD class="tdheaderline" align="center" width="17px">11</TD><TD class="tdheaderline" align="center" width="17px">12</TD><TD class="tdheaderline" align="center" width="17px">13</TD><TD class="tdheaderline" align="center" width="17px">14</TD><TD class="tdheaderline" align="center" width="17px">15</TD><TD class="tdheaderline" align="center" width="17px">16</TD><TD class="tdheaderline" align="center" width="17px">17</TD><TD class="tdheaderline" align="center" width="17px">18</TD><TD class="tdheaderline" align="center" width="17px">19</TD><TD class="tdheaderline" align="center" width="17px">20</TD><TD class="tdheaderline" align="center" width="17px">21</TD><TD class="tdheaderline" align="center" width="17px">22</TD><TD class="tdheaderline" align="center" width="17px">23</TD><TD class="tdheaderline" align="center" width="17px">24</TD>

						
					</TR>
					
					<TR>
					<TR><TD class=form align=center>Infantry</TD><TD class=form align=center></TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR><TR><TD class=form align=center>Commandoes</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center></TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR><TR><TD class=form align=center>Jeeps</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR><TR><TD class=form align=center>Matilda MkI</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR><TR><TD class=form align=center>Matilda MkII</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR><TR><TD class=form align=center>Sherman</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR><TR><TD class=form align=center>Spitfire MkI</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR><TR><TD class=form align=center>Spitfire MkII</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR><TR><TD class=form align=center>Spitfire MkV</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR><TR><TD class=form align=center>Hurricance MkI</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR><TR><TD class=form align=center>Halifax Bomber</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR><TR><TD class=form align=center>Wellington Bomber</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR><TR><TD class=form align=center>Lancaster Bomber</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR><TR><TD class=form align=center>Landing Craft</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR><TR><TD class=form align=center>Landing Ship</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR><TR><TD class=form align=center>Frigates</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR><TR><TD class=form align=center>Battleships</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR><TR><TD class=form align=center>Aircraft Carriers</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD><TD class=form align=center>&nbsp;</TD></TR>

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