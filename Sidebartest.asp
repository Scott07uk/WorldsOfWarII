<!---#include file="includes\check.asp"--->
<!---SPELL CHECKED--->

<HTML>
	<HEAD>
		<TITLE>Sidebar</title>
		<link rel=stylesheet HREF="includes/all.css">
		<SCRIPT language="javascript">
			//<--hide script from older browsers
			//function to open a new window
			function openWin(url)
			{
				//open the new window
				window.open(url,'irclcient', config='height=450,width=650,toolbar=no');
			}	
			//--end hiding>
		</script>
	</HEAD>
	<%
	Dim rsconfig
	strsql = "SELECT portaladdress, forumsaddress, manualaddress, ircsite FROM tblconfig;"
	
	set rsconfig = server.CreateObject("ADODB.Recordset")
	
	rsconfig.Open strsql, adocon
	%>
	<Body bgcolor="000000" text="#ffffff" link="#ffffff" alink="#ffffff" vlink="#ffffff" leftmargin="0" topmargin="0" marginleft="0" margintop="0">
		<table width="180px" background="images\logo.gif" cellspacing="0" cellpadding="0" style="background-repeat: no-repeat; border: none;">
			<tr>
				<td height="138px">
				</td>
			</tr>
			<tr>
				<td height="5px" NOWRAP>
				</td>
			</tr>
			<tr>
				<td>
					<img src="images/1.gif" border="0" Usemap="#1">
				</td>
			</tr>
			<tr>
				<td height="5px" NOWRAP>
				</td>
			</tr>
			<tr>
				<td>
					<img src="images/2.gif" border="0" Usemap="#2">
				</td>
			</tr>
			<tr>
				<td height="5px" NOWRAP>
				</td>
			</tr>
			<tr>
				<td>
					<img src="images/3.gif" border="0" Usemap="#3">
				</td>
			</tr>
			<tr>
				<td height="5px" NOWRAP>
				</td>
			</tr>
			<tr>
				<td>
					<img src="images/4.gif" border="0" Usemap="#4">
				</td>
			</tr>
			<tr>
				<td height="5px" NOWRAP>
				</td>
			</tr>
			<tr>
				<td>
					<img src="images/5.gif" border="0" Usemap="#5">
				</td>
			</tr>
			<MAP NAME="1">
				<AREA SHAPE=RECT COORDS="0,0,169,19" HREF="over.asp"  ALT="Overview"  TARGET="B">
				<AREA SHAPE=RECT COORDS="0,20,169,37" HREF="priv_news.asp"  ALT="Private news"  TARGET="B">
				<AREA SHAPE=RECT COORDS="0,38,169,55" HREF="continent_status.asp"  ALT="Continent Status"  TARGET="B">
				<AREA SHAPE=RECT COORDS="0,56,169,74" HREF="msgs.asp"  ALT="Mail"  TARGET="B">
				<AREA SHAPE=RECT COORDS="0,75,169,91" HREF="msgboard.asp"  ALT="Message Board"  TARGET="B">
				<AREA SHAPE=RECT COORDS="0,92,169,114" HREF="javascript:openWin('wow_irc.asp');"  ALT="IRC"  TARGET="B">
			</MAP>
			<MAP NAME="2">
				<AREA SHAPE=RECT COORDS="0,0,169,19" HREF="explore.asp"  ALT="Exploration"  TARGET="B">
				<AREA SHAPE=RECT COORDS="0,20,169,37" HREF="development.asp"  ALT="Development"  TARGET="B">
				<AREA SHAPE=RECT COORDS="0,38,169,55" HREF="research.asp"  ALT="Research"  TARGET="B">
				<AREA SHAPE=RECT COORDS="0,56,169,72" HREF="construction.asp"  ALT="Construction"  TARGET="B">
			</MAP>
			<MAP NAME="3">
				<AREA SHAPE=RECT COORDS="0,0,169,19" HREF="units.asp"  ALT="Military Units"  TARGET="B">
				<AREA SHAPE=RECT COORDS="0,20,169,37" HREF="defence.asp"  ALT="Defence Network"  TARGET="B">
				<AREA SHAPE=RECT COORDS="0,38,169,55" HREF="mil_opps.asp"  ALT="Military Operations"  TARGET="B">
				<AREA SHAPE=RECT COORDS="0,56,169,74" HREF="spies.asp"  ALT="Special Operations Executive"  TARGET="B">
			</MAP>
			<MAP NAME="4">
				<AREA SHAPE=RECT COORDS="0,0,169,19" HREF="politics.asp"  ALT="Politics"  TARGET="B">
				<AREA SHAPE=RECT COORDS="0,20,169,37" HREF="continent.asp"  ALT="Continent"  TARGET="B">
				<AREA SHAPE=RECT COORDS="0,38,169,55" HREF="world.asp"  ALT="World"  TARGET="B">
				<AREA SHAPE=RECT COORDS="0,56,169,74" HREF="settings.asp"  ALT="Settings"  TARGET="B">
			</MAP>
			<MAP NAME="5">
				<AREA SHAPE=RECT COORDS="0,0,169,19" HREF="logout.asp"  ALT="Logout"  TARGET="_top">
				<AREA SHAPE=RECT COORDS="0,20,169,37" HREF="<% =rsconfig("ircsite") %>"  ALT="IRC Website"  TARGET="_blank">
				<AREA SHAPE=RECT COORDS="0,38,169,55" HREF="<% =rsconfig("forumsaddress") %>"  ALT="Forums"  TARGET="_blank">
				<AREA SHAPE=RECT COORDS="0,56,169,74" HREF="<% =rsconfig("portaladdress") %>"  ALT="Portal"  TARGET="_top">
				<AREA SHAPE=RECT COORDS="0,75,169,91" HREF="<% =rsconfig("manualaddress") %>"  ALT="Manual"  TARGET="_blank">
			</MAP>			
		
		</table>
	</body>
	<%
	rsconfig.Close 
	
	set rsconfig = nothing
	%>
</html>
<!---#include file="includes\end_check.asp"--->