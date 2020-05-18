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
<%								
if blnAdsOn = true then
Dim RsAd
								
set rsad = server.CreateObject("ADODB.Recordset")
		strsql = "SELECT adText, adshows FROM tblBanners WHERE Active = 1 ORDER by adshows ASC;"
										
		rsad.LockType = 3 
		rsad.CursorType = 2
										
		rsad.open strsql, adocon
		if not rsad.EOF then
				call rsad.Update("adshows", clng(rsad("adshows"))+1)
%>
<center>
	<% = rsad("adtext") %>
</center>
<BR />
<%
		end if
										
		rsad.Close 
		set rsad = nothing							
end if
%>

<table width="95%" cellpadding="0" cellspacing="0" align="center" style="border-top: solid; border-left: solid; border-bottom: solid; border-right: solid; border-width: 1;">
	<tr>
		<td>
			<table width="100%" cellpadding="1" cellspacing="0" class="norm">
				<tr>
					<td class="tdheaderline" align="center" colspan="7">
						<%=strLeadername%> of <%=strCountryname%> [<%=lngContinentno%>:<%=lngCountryno%>]
					</td>
				</tr>
				<tr>
					<td style="border-right: solid; border-bottom: solid; border-width: 1;" width="14%">
						Food
						<br />
						<%=formatnumber(lngfood,0)%>
					</td>
					<td style="border-right: solid; border-bottom: solid; border-width: 1;" width="14%">
						Wood
						<br />
						<%=formatnumber(lngWood,0)%>
					</td>
					<td style="border-right: solid; border-bottom: solid; border-width: 1;" width="15%">
						Iron
						<br />
						<%=formatnumber(lngIron,0)%>
					</td>
					<td style="border-right: solid; border-bottom: solid; border-width: 1;" width="14%">
						Money
						<br />
						<%=formatnumber(lngMoney,0)%>
					</td>
					<td style="border-right: solid; border-bottom: solid; border-width: 1;" width="15%">
						Land
						<br />
						<%=formatnumber(lngLand,0)%>
					</td>
					<td style="border-right:solid; border-bottom: solid; border-width: 1;" width="14%">
						Gold
						<br />
						<%=formatnumber(lnggold,0)%>
					</td>
					<TD style="border-bottom:solid; border-width:1;" width="14%">
						Score
						<BR />
						<% =formatnumber(lngScore,0) %>
					</TD>
				</tr>
				<tr>
					<td align="center" colspan="7">
						<table width="100%" cellspacing="0" cellpadding="0">
							<tr>
								<td align="left" width="30%">
<%
Dim rsMessages
Dim blnNewMsg
Dim blnNewNews
Dim blnInAtt
Dim blnInDef
set rsMessages = server.CreateObject("ADODB.Recordset")

strSql = "SELECT tblmessages.new FROM tblmessages WHERE tocountry = " & lngCountryno & "AND tocontinent = " & lngContinentNo & "AND new = 1;"

rsmessages.Open strSQL,adocon

if rsmessages.EOF then
	blnnewmsg = False
else
	blnnewmsg = True
end if

rsMessages.Close

strSql = "SELECT tblnews.unread FROM tblnews WHERE username = '" & strusername & "'AND unread = 1;"

rsMessages.Open strSQL,adocon

if rsMessages.EOF then
	blnnewnews = false
else
	blnnewnews = true
end if

rsmessages.close

strsql = "SELECT status FROM tblmovements WHERE tocontinent = " & lngcontinentno & " AND tocountry = " & lngcountryno

rsMessages.Open strsql, adocon
blninatt = false
blnindef = false
do while not rsMessages.EOF 
	if rsmessages("status") = 1 then
			blninatt = true
	elseif rsmessages("status")= 2 then
			blnindef = true
	end if
	rsMessages.MoveNext 
loop

rsMessages.Close 

set rsmessages = NOTHING
%>
									<a href="msgs.asp">
<% if blnnewmsg = TRUE then %>
									<img src="images/mailon.gif" border="0">
<% else %>
									<img src="images/mailoff.gif" border="0">
<% end if %>
									</a> | <a href="priv_news.asp">
<% if blnnewnews = true then %>
									<img src="images/newson.gif" border="0">
<% else %>
									<img src="images/newsoff.gif" border="0">
<% end if %>
									</a>
								</td>
								
								</TD>
								<TD width="40%" align="center">
								<% if blninatt = true then %>
								<font color="#FF0000">
									INCOMING ATTACKERS
								</FONT>
								<%
								end if
								if blnindef = true then 
								%>
								<font color="#00FF00">
									&nbsp;&nbsp;INCOMING DEFENDERS
								</FONT>
								<% END IF %>
								</TD>
								<td width="30%" align="right" NOWRAP>
									<%
									if blnTickRunning = True then
											Response.Write("Tick Running")
										else
									strsql = "SELECT tblticktimes.ticktime FROM tblticktimes WHERE ticktime > " & minute(now())
									Dim rsticks
									
									set rsticks = server.CreateObject("ADODB.Recordset")
									
									rsticks.Open strsql, adocon
									
									if rsticks.EOF then
											strsql = "SELECT tblticktimes.ticktime FROM tblticktimes WHERE ticktime > " & cint(minute(now()))-60
											
											rsticks.Close 
											
											rsticks.Open strsql, adocon
											if rsticks.EOF then
									%>
									Ticks Stopped
									<%
									else
									%>
									Next Tick: <%=cint(rsticks("ticktime") - minute(now())+60)%> Minutes
									<%
									end if
									else
									%>
									Next Tick: <%=cint(rsticks("ticktime") - minute(now()))%> Minutes
									<%
									end if
									
									rsticks.Close
									
									set rsticks = nothing
									end if
									%>
								</td>
							</tr>
						</table>
					</td>
				</tr>								
			</table>
		</td>
	</tr>
</table>