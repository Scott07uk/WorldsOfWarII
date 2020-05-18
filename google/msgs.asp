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
			<A HREF="javascript:openHelp('5');">Help With This Page</A>
		</TD>
	</TR>
</table>
</center>


<Script language="javascript">
function selectAll() {
for (counter=0;counter<document.deletemsg.box.length;counter++)
{
document.deletemsg.box[counter].checked=document.deletemsg.selectall.checked;
}
}

</Script>
<br />
<center>

<form action="deletemsg_engine.asp" method="post" name="deletemsg" ID="Form1">
<table style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;" cellpadding="5" cellspacing="5" width="95%" ID="Table5">
<tr>
<td class="tdheader" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;" width="20"><input onclick="javascript:selectAll();" type="checkbox" name="selectall" ID="Checkbox1"></td>
<td class="tdheader" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;" width="100"><b>

<a href="msgs.asp?order=1">Co-ords</a>

</b></td>
<td class="tdheader" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;" width="200"><b>

<a href="msgs.asp?order=2">Date Received</a>

</b></td>
<td class="tdheader" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;"><b>

<a href="msgs.asp?order=3">Subject</a>

</b></td>
</tr>

</table>
No Messages

<br />
<br />
<table style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;" cellpadding="1" cellspacing="5" width="95%" ID="Table6">
<td width="60"><input type="submit" value="Delete" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1; padding:0;" class="form" ID="Submit1" NAME="Submit1"></td>
</form>

<td><form action="sendmsg.asp" method="post" ID="Form2"></td>
<td width="60"><input type="submit" value="New Message" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1; padding:0" align="right" width="120" class="form" ID="Submit2" NAME="Submit2"></td>
</form>
</table>
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