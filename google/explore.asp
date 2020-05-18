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


<table width="95%" cellpadding="0" cellspacing="0" align="center" style="border-top: solid; border-left: solid; border-bottom: solid; border-right: solid; border-width: 1;" ID="Table1">
	<tr>
		<td>

			<table width="100%" cellpadding="1" cellspacing="0" class="norm" ID="Table2">
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
						<table width="100%" cellspacing="0" cellpadding="0" ID="Table3">
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

<table width="95%" cellpadding="0" border="0" cellspacing="0" style="border-top:none; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;" ID="Table4">
	<TR>
		<TD align="center" class="norm">
			<A HREF="javascript:openHelp('8');">Help With This Page</A>
		</TD>
	</TR>
</table>
</center>


<script language="javascript">
//<--being hiding

function checkOrder(){
//function to make sure the ordering new explorers form is correct
var rtnVar;
if (document.explore.explorers.value==""){
	rtnVar=false;
	alert('You have not entered how many units you would like to order.');
}
else{
	rtnVar=true;
}

return rtnVar;
}

function calCost(){
var initialCost=0;
var curLand=0;
var landWanted = document.exploreland.acres.value;
var cost=0;

initialCost=curLand;
cost=initialCost;

if (landWanted<1){
document.exploreland.display.value = "You cannot explore less than 1 acre."
}
else{

for (i=1; i<landWanted; i++){
	cost=cost+initialCost+i;
}
document.exploreland.display.value="Exploring "+document.exploreland.acres.value+" acres will require "+cost+" explorers.";
}
return true;
}


//-->

</script>
<BR />
<BR />
<center>

	<table cellpadding="2" width="95%" cellspacing="0" border="0" style="border-top:solid; border-bottom:solid; border-right:solid; border-left:solid; border-color:#FFFFFF; border-width:1;" ID="Table5">
		<TR>
			<TD class="tdheaderline" align="center">
				Training Explorers
			</TD>
		</TR>
		<TR>

			<TD class="norm" align="center">
				
				You have no explorers currently in training.
				
			</TD>
		</TR>
	</table>
	<BR />
	<BR />
	<table width="95%" cellpadding="2" cellspacing="0" border="0" style="border-top:solid; border-bottom:solid; border-right:solid; border-left:solid; border-color:#FFFFFF; border-width:1;" ID="Table6">
		<TR>

			<TD class="tdheaderline" align="center">
				Train New Explorers
			</TD>
		</TR>
		<TR>
			<TD class="norm" align="center">
				<form action="order_explore.asp" method="post" name="explore" onsubmit="return checkOrder();" ID="Form1">
					<table cellpadding="5" border="0" width="100%" cellspacing="0" ID="Table7">
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
								Order
							</TD>

						</TR>
						<TR title="Explorers allow you to explore and take control of more land, allowing you to build more industry.">
							<TD class="norm">
								Explorers
							</TD>
							<TD class="norm">
								250
							</TD>
							<TD class="norm">
								10
							</TD>

							<TD class="norm">
								50
							</TD>
							<TD class="norm">
								500
							</TD>
							<TD class="norm">
								6 ticks
							</TD>
							<TD class="norm">

								<input type="text" size="2" name="explorers" class="form" maxlength="3" ID="Text1"/>
							</TD>
						</TR>
					</table>
					<BR />
					<input type="submit" value="Order" class="form" ID="Submit1" NAME="Submit1">
				</form>
			</TD>
		</TR>

	</table>
	<BR />
	<BR />
	<table width="95%" cellpadding="2" cellspacing="0" border="0" style="border-top:solid; border-bottom:solid; border-right:solid; border-left:solid; border-color:#FFFFFF; border-width:1;" ID="Table8">
		<TR>
			<TD class="tdheaderline" align="center">
				Explore New Land
			</TD>
		</TR>

		<TR>
			<TD class="norm" align="center">
				The greater the amount of land you own, the more explorers you will need to bring more under your control.
				<BR />
				You currently have 0 you can use to explore new land.  
				<Br />
				<Br />
				<form action="explore_land.asp" method="post" name="exploreland" ID="Form2">
					Number of acres to explore:
					<input type="text" size="2" class="form" name="acres" maxlength="3" value="0" ID="Text2"/>

					<BR />
					<BR />
					<input type="button" class="form" name="Cal" value="Calculate" onclick="calCost();" ID="Button1"/>
					<input type="text" size="60" name="display" class="form" readonly ID="Text3"/>
					<BR />
					<BR />
					<input type="submit" value="Explore" class="form" ID="Submit2" NAME="Submit2"/>
				</form>
			</TD>

		</TR>
	</table>
</center>
<BR />

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