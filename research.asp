<!---#include file="includes/check.asp"--->
<!---#include file="includes/top.asp"--->

<!---SPELL CHECKED--->
<center>
<table width="95%" cellpadding="0" border="0" cellspacing="0" style="border-top:none; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD align="center" class="norm">
			<A HREF="javascript:openHelp('10');">Help With This Page</A>
		</TD>
	</TR>
</table>
<BR />
<BR />
<%
select case request.querystring("error")
	case 1
		Response.Write("The tick is currently processing, you can not start a new research during that time, please try again in a few seconds.  <BR /><BR />")
end select
%>

</center>

<%
Set rsResearch = Server.CreateObject("ADODB.Recordset")
		
strSQL = "SELECT tblaccounts.username, researchticks, resSF, resCE, resVGM, resBGP, resMPV, resMWP, resMPC, resFPC, resWC, resRP, resUWR, resG, resEP, resLB, resSIT, resBLC, resSBS, resMBS, resAAS, resAES, resCS, resFG, resWMP, resE, resD, resAC  FROM tblaccounts WHERE username = '" & strUsername & "';"
		
rsResearch.Open strSQL, adoCon
		
if not rsResearch.EOF then

if rsResearch("resCE") = 2 then
	resprogress = 1
	res1 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resCE") = 3 then
	res1 = 1
else
	res1 = 0
end if

if rsResearch("resVGM") = 2 then
	resprogress = 2
	res2 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resVGM") = 3 then
	res2 = 1
else
	res2 = 0
end if

if rsResearch("resSF") = 2 then
	resprogress = 3
	res3 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resSF") = 3 then
	res3 = 1
else
	res3 = 0
end if

if rsResearch("resMPV") = 2 then
	resprogress = 4
	res4 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resMPV") = 3 then
	res4 = 1
else
	res4 = 0
end if

if rsResearch("resBGP") = 2 then
	resprogress = 5
	res5 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resBGP") = 3 then
	res5 = 1
else
	res5 = 0
end if

if rsResearch("resMWP") = 2 then
	resprogress = 6
	res6 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resMWP") = 3 then
	res6 = 1
else
	res6 = 0
end if

if rsResearch("resMPC") = 2 then
	resprogress = 7
	res7 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resMPC") = 3 then
	res7 = 1
else
	res7 = 0
end if

if rsResearch("resWC") = 2 then
	resprogress = 8
	res8 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resWC") = 3 then
	res8 = 1
else
	res8 = 0
end if

if rsResearch("resRP") = 2 then
	resprogress = 9
	res9 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resRP") = 3 then
	res9 = 1
else
	res9 = 0
end if

if rsResearch("resUWR") = 2 then
	resprogress = 10
	res10 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resUWR") = 3 then
	res10 = 1
else
	res10 = 0
end if

if rsResearch("resFPC") = 2 then
	resprogress = 11
	res11 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resFPC") = 3 then
	res11 = 1
else
	res11 = 0
end if

if rsResearch("resEP") = 2 then
	resprogress = 12
	res12 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resEP") = 3 then
	res12 = 1
else
	res12 = 0
end if

if rsResearch("resLB") = 2 then
	resprogress = 13
	res13 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resLB") = 3 then
	res13 = 1
else
	res13 = 0
end if

if rsResearch("resG") = 2 then
	resprogress = 14
	res14 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resG") = 3 then
	res14 = 1
else
	res14 = 0
end if

if rsResearch("resSIT") = 2 then
	resprogress = 15
	res15 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resSIT") = 3 then
	res15 = 1
else
	res15 = 0
end if

if rsResearch("resBLC") = 2 then
	resprogress = 16
	res16 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resBLC") = 3 then
	res16 = 1
else
	res16 = 0
end if

if rsResearch("resSBS") = 2 then
	resprogress = 17
	res17 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resSBS") = 3 then
	res17 = 1
else
	res17 = 0
end if

if rsResearch("resMBS") = 2 then
	resprogress = 18
	res18 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resMBS") = 3 then
	res18 = 1
else
	res18 = 0
end if

if rsResearch("resAAS") = 2 then
	resprogress = 19
	res19 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resAAS") = 3 then
	res19 = 1
else
	res19 = 0
end if

if rsResearch("resFG") = 2 then
	resprogress = 20
	res20 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resFG") = 3 then
	res20 = 1
else
	res20 = 0
end if

if rsResearch("resCS") = 2 then
	resprogress = 21
	res21 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resCS") = 3 then
	res21 = 1
else
	res21 = 0
end if

if rsResearch("resWMP") = 2 then
	resprogress = 22
	res22 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resWMP") = 3 then
	res22 = 1
else
	res22 = 0
end if

if rsResearch("resE") = 2 then
	resprogress = 23
	res23 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resE") = 3 then
	res23 = 1
else
	res23 = 0
end if

if rsResearch("resD") = 2 then
	resprogress = 24
	res24 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resD") = 3 then
	res24 = 1
else
	res24 = 0
end if

if rsResearch("resAC") = 2 then
	resprogress = 25
	res25 = 0
	restime = rsResearch("researchticks")
elseif rsResearch("resAC") = 3 then
	res25 = 1
else
	res25 = 0
end if
end if
rsResearch.close
set rsResearch = nothing
%>

<!---#include file="includes/techtree.asp"--->

<BR />
<table align="center" class="norm" cellpadding="2" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<tr>
		<TD class="tdheaderline" width="">
			Research
		</TD>
		<td class="tdheaderline" width="">
			Food
		</td>
		<td class="tdheaderline" width="">
			Wood
		</td>
		<td class="tdheaderline" width="">
			Iron
		</td>
		<td class="tdheaderline" width="">
			Money
		</td>
		<td class="tdheaderline" width="">
			Time
		</td>
		<td class="tdheaderline" width="">&nbsp;
			
		</td>
	</TR>
	<TR height="27"><!---#combustion engines for military use--->
		<td class="all" valign="top">
			Combustion Engines for Military Use
		</td>
		<td class="all">
			500
		</td>
		<td class="all">
			0
		</td>
		<td class="all">
			1000
		</td>
		<td class="all">
			1250
		</td>
		<td class="all" valign="top">
			10 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 1 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res1 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res1 = 1 then 
%>
			Research Complete
<% elseif res1 = 0 then %>
			<a href="Research_engine.asp?ID=1">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This allows the research and development of mobile systems that carry more powerful guns over greater distances, such as tanks and planes.</i></small>
		</td>
	</tr>
<% 
if available2 = 1 then 
%>
	<tr height="27"><!---#Mobile gun mounting--->
		<td class="all" valign="top">
			Mobile Gun Mounting
		</td>
		<td class="all">
			500
		</td>
		<td class="all">
			0
		</td>
		<td class="all">
			850
		</td>
		<td class="all">
			1000
		</td>
		<td class="all" valign="top">
			12 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 2 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res2 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res2 = 1 then 
%>
			Research Complete
<% elseif res2 = 0 then %>
			<a href="Research_engine.asp?ID=2">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This allows larger guns to be mounted on mobile units.  This includes mounting Browning 303 machine guns on jeeps, tanks and planes.</i></small>
		</td>
	</tr>
<% 
end if 
%>
	<tr height="27"><!---#Special forces--->
		<td class="all" valign="top">
			Special Forces
		</td>
		<td class="all">
			500
		</td>
		<td class="all">
			200
		</td>
		<td class="all">
			500
		</td>
		<td class="all">
			2000
		</td>
		<td class="all" valign="top">
			8 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 3 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res3 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res3 = 1 then 
%>
			Research Complete
<% elseif res3 = 0 then %>
			<a href="Research_engine.asp?ID=3">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>Special forces allow more advanced troops to be trained for both military operations and more underhand tactics.</i></small>
		</td>
	</tr>
<% 
if available4 = 1 then 
%>
	<tr height="27"><!---#mass production of military vehicles--->
		<td class="all" valign="top">
			Mass Production of Military Vehicles
		</td>
		<td class="all">
			700
		</td>
		<td class="all">
			500
		</td>
		<td class="all">
			2000
		</td>
		<td class="all">
			2000
		</td>
		<td class="all" valign="top">
			15 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 4 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res4 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res4 = 1 then 
%>
			Research Complete
<% elseif res4 = 0 then %>
			<a href="Research_engine.asp?ID=4">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This allows the production of military vehicles such as tanks on a scale that is usable by the army.</i></small>
		</td>
	</tr>
<% 
end if
if available5 = 1 then 
%>
	<tr height="27"><!---#balancing large guns--->
		<td class="all" valign="top">
			Balancing Large Guns
		</td>
		<td class="all">
			700
		</td>
		<td class="all">
			0
		</td>
		<td class="all">
			1500
		</td>
		<td class="all">
			1750
		</td>
		<td class="all" valign="top">
			12 ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 5 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res5 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res5 = 1 then 
%>
			Research Complete
<% elseif res5 = 0 then %>
			<a href="Research_engine.asp?ID=5">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>Being able to balance large guns allows you to make mobile large guns, such as the ones used in the larger tanks.</i></small>
		</td>
	</tr>
	
<% 
end if
if available6 = 1 then 
%>
	<tr height="27"><!---#monowing planes--->
		<td class="all" valign="top">
			Mono-Wing Planes
		</td>
		<td class="all">
			500
		</td>
		<td class="all">
			500
		</td>
		<td class="all">
			1000
		</td>
		<td class="all">
			2500
		</td>
		<td class="all" valign="top">
			18 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 6 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res6 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res6 = 1 then 
%>
			Research Complete
<% elseif res6 = 0 then %>
			<a href="Research_engine.asp?ID=6">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>All modern fighters and bombers are mono-wing planes. To be able to build any you must research this.</i></small>
		</td>
	</tr>
<% 
end if
if available7 = 1 then 
%>
	<tr height="27">
		<td class="all" valign="top">
			All Metal Plane Construction
		</td>
		<td class="all">
			750
		</td>
		<td class="all">
			0
		</td>
		<td class="all">
			1500
		</td>
		<td class="all">
			2500
		</td>
		<td class="all" valign="top">
			20 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 7 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res7 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res7 = 1 then 
%>
			Research Complete
<% elseif res7 = 0 then %>
			<a href="Research_engine.asp?ID=7">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This researches the use of metal components to manufacture the new breed of fighters and bombers.</i></small>
		</td>
	</tr>
<% 
end if
if available8 = 1 then 
%>
	<tr height="27">
		<td class="all" valign="top">
			Wing Cannons
		</td>
		<td class="all">
			1000
		</td>
		<td class="all">
			0
		</td>
		<td class="all">
			250
		</td>
		<td class="all">
			2500
		</td>
		<td class="all" valign="top">
			15 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 8 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res8 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res8 = 1 then 
%>
			Research Complete
<% elseif res8 = 0 then %>
			<a href="Research_engine.asp?ID=8">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This puts in to use the idea of putting small armed shells in fighter planes so they might have more destructive power.</i></small>
		</td>
	</tr>
<% 
end if
if available9 = 1 then 
%>
	<tr height="27">
		<td class="all" valign="top">
			Rocket Propulsion
		</td>
		<td class="all">
			1500
		</td>
		<td class="all">
			100
		</td>
		<td class="all">
			750
		</td>
		<td class="all">
			2750
		</td>
		<td class="all" valign="top">
			24 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 9 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res9 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res9 = 1 then 
%>
			Research Complete
<% elseif res9 = 0 then %>
			<a href="Research_engine.asp?ID=9">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>A new idea being developed, having some fully propelled shell for long distance strafing and disemboweling bombers.</i></small>
		</td>
	</tr>
<% 
end if
if available10 = 1 then 
%>
	<tr height="27">
		<td class="all" valign="top">
			Under Wing Rockets
		</td>
		<td class="all">
			1250
		</td>
		<td class="all">
			0
		</td>
		<td class="all">
			750
		</td>
		<td class="all">
			2500
		</td>
		<td class="all" valign="top">
			18 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 10 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res10 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res10 = 1 then 
%>
			Research Complete
<% elseif res10 = 0 then %>
			<a href="Research_engine.asp?ID=10">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This develops the methods used to put rockets under wings in a way that won't affect the planes performance.</i></small>
		</td>
	</tr>
<% 
end if
if available11 = 1 then 
%>
	<tr height="27">
		<td class="all" valign="top">
			Fabric Plane Covers
		</td>
		<td class="all">
			750
		</td>
		<td class="all">
			1500
		</td>
		<td class="all">
			0
		</td>
		<td class="all">
			2000
		</td>
		<td class="all" valign="top">
			20 ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 11 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res11 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res11 = 1 then 
%>
			Research Complete
<% elseif res11 = 0 then %>
			<a href="Research_engine.asp?ID=11">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This develops the idea that planes not fully made of metal are much simpler to repair and it can be done in the field.</i></small>
		</td>
	</tr>
<% 
end if
if available12 = 1 then 
%>
	<tr height="27">
		<td class="all" valign="top">
			Efficient Payloads
		</td>
		<td class="all">
			1500
		</td>
		<td class="all">
			1000
		</td>
		<td class="all">
			1000
		</td>
		<td class="all">
			1750
		</td>
		<td class="all" valign="top">
			24 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 12 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res12 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res12 = 1 then 
%>
			Research Complete
<% elseif res12 = 0 then %>
			<a href="Research_engine.asp?ID=12">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This is the first stage of developing a way of holding a large explosive charge to be dropped from a new bomber.</i></small>
		</td>
	</tr>
<% 
end if
if available13 = 1 then 
%>
	<tr height="27">
		<td class="all" valign="top">
			Larger Bombs
		</td>
		<td class="all">
			1000
		</td>
		<td class="all">
			500
		</td>
		<td class="all">
			750
		</td>
		<td class="all">
			1250
		</td>
		<td class="all" valign="top">
			16 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 13 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res13 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res13 = 1 then 
%>
			Research Complete
<% elseif res13 = 0 then %>
			<a href="Research_engine.asp?ID=13">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This expands on existing technology to allow smaller bombers pack a bigger punch.</i></small>
		</td>
	</tr>
<% 
end if
if available14 = 1 then 
%>
	<tr height="27">
		<td class="all" valign="top">
			G
		</td>
		<td class="all">
			1000
		</td>
		<td class="all">
			250
		</td>
		<td class="all">
			500
		</td>
		<td class="all">
			3750
		</td>
		<td class="all" valign="top">
			36 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 14 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res14 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res14 = 1 then 
%>
			Research Complete
<% elseif res14 = 0 then %>
			<a href="Research_engine.asp?ID=14">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>The G navigation system will allow accurate long range bombing at night in the largest bombers.</i></small>
		</td>
	</tr>
<% 
end if
%>
	<tr height="27">
		<td class="all" valign="top">
			Sea Invasion Tactics
		</td>
		<td class="all">
			750
		</td>
		<td class="all">
			500
		</td>
		<td class="all">
			500
		</td>
		<td class="all">
			1250
		</td>
		<td class="all" valign="top">
			8 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 15 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res15 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res15 = 1 then 
%>
			Research Complete
<% elseif res15 = 0 then %>
			<a href="Research_engine.asp?ID=15">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>Works on planning the best and most useful ways to land troops on foreign shores.</i></small>
		</td>
	</tr>
<% 
if available16 = 1 then 
%>
	<tr height="27">
		<td class="all" valign="top">
			Improved Landing Craft
		</td>
		<td class="all">
			1250
		</td>
		<td class="all">
			1000
		</td>
		<td class="all">
			1500
		</td>
		<td class="all">
			2000
		</td>
		<td class="all" valign="top">
			12 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 16 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res16 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res16 = 1 then 
%>
			Research Complete
<% elseif res16 = 0 then %>
			<a href="Research_engine.asp?ID=16">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This develops on existing ideas but allows tanks and other mechanised units to be landed on foreign shores.</i></small>
		</td>
	</tr>
<% 
end if
%>
	<tr height="27">
		<td class="all" valign="top">
			Modern Battleships
		</td>
		<td class="all">
			3250
		</td>
		<td class="all">
			2500
		</td>
		<td class="all">
			5500
		</td>
		<td class="all">
			4750
		</td>
		<td class="all" valign="top">
			48 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 17 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res17 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res17 = 1 then 
%>
			Research Complete
<% elseif res17 = 0 then %>
			<a href="Research_engine.asp?ID=17">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This brings the naval units into the modern age, so they can defend our shores.</i></small>
		</td>
	</tr>
<% 
if available18 = 1 then 
%>
	<tr height="27">
		<td class="all" valign="top">
			Improved Modern Battleships
		</td>
		<td class="all">
			4500
		</td>
		<td class="all">
			3000
		</td>
		<td class="all">
			6750
		</td>
		<td class="all">
			5500
		</td>
		<td class="all" valign="top">
			36 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 18 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res18 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res18 = 1 then 
%>
			Research Complete
<% elseif res18 = 0 then %>
			<a href="Research_engine.asp?ID=18">Start</a>		
<% end if %>
		</td>
	</TR>	
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This makes the designs of existing battleships bigger and better, more guns more armour and more speed.</i></small>
		</td>
	</tr>
<% 
end if
if available19 = 1 then 
%>
	<tr height="27">
		<td class="all" valign="top">
			Aviation at Sea
		</td>
		<td class="all">
			6750
		</td>
		<td class="all">
			3500
		</td>
		<td class="all">
			8250
		</td>
		<td class="all">
			8250
		</td>
		<td class="all" valign="top">
			65 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 19 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res19 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res19 = 1 then 
%>
			Research Complete
<% elseif res19 = 0 then %>
			<a href="Research_engine.asp?ID=19">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>It's been proved that planes have power over sea, so why not give them extra range by putting planes on boats?</i></small>
		</td>
	</tr>
<% 
end if
%>
	<tr height="27">
		<td class="all" valign="top">
			Fixed Gun Positions
		</td>
		<td class="all">
			750
		</td>
		<td class="all">
			500
		</td>
		<td class="all">
			1500
		</td>
		<td class="all">
			1250
		</td>
		<td class="all" valign="top">
			15 ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 20 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res20 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res20 = 1 then 
%>
			Research Complete
<% elseif res20 = 0 then %>
			<a href="Research_engine.asp?ID=20">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This develops the best stragety to make guns fixed to the ground.</i></small>
		</td>
	</tr>
	<tr height="27">
		<td class="all" valign="top">
			Chemical Usage
		</td>
		<td class="all">
			1000
		</td>
		<td class="all">
			0
		</td>
		<td class="all">
			250
		</td>
		<td class="all">
			1250
		</td>
		<td class="all" valign="top">
			18 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 21 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res21 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res21 = 1 then 
%>
			Research Complete
<% elseif res21 = 0 then %>
			<a href="Research_engine.asp?ID=21">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This develops ways of extracting chemicals which are of use to the military on an industrial scale.</i></small>
		</td>
	</tr>
	<tr height="27">
		<td class="all" valign="top">
			Water Mine Patterns
		</td>
		<td class="all">
			750
		</td>
		<td class="all">
			1250
		</td>
		<td class="all">
			1750
		</td>
		<td class="all">
			1500
		</td>
		<td class="all" valign="top">
			18 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 22 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res22 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res22 = 1 then 
%>
			Research Complete
<% elseif res22 = 0 then %>
			<a href="Research_engine.asp?ID=22">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This develops a strategy to allow the use of water mines while allowing our ships to pass.</i></small>
		</td>
	</tr>
	<tr height="27">
		<td class="all" valign="top">
			Encryption
		</td>
		<td class="all">
			1750
		</td>
		<td class="all">
			50
		</td>
		<td class="all">
			150
		</td>
		<td class="all">
			5750
		</td>
		<td class="all" valign="top">
			32 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 23 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res23 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res23 = 1 then 
%>
			Research Complete
<% elseif reS23 = 0 then %>
			<a href="Research_engine.asp?ID=23">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This develops ways of hiding information from the enemy, mainly to be used by SEO.</i></small>
		</td>
	</tr>
	<tr height="27">
		<td class="all" valign="top">
			Deception
		</td>
		<td class="all">
			1750
		</td>
		<td class="all">
			500
		</td>
		<td class="all">
			1500
		</td>
		<td class="all">
			2500
		</td>
		<td class="all" valign="top">
			16 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 24 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res24 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res24 = 1 then 
%>
			Research Complete
<% elseif reS24 = 0 then %>
			<a href="Research_engine.asp?ID=24">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>Develops some of the less government friendly tricks used by SEO.</i></small>
		</td>
	</tr>
<%
if available25 = 1 then
%>
	<tr height="27">
		<td class="all" valign="top">
			Advanced Communications
		</td>
		<td class="all">
			2000
		</td>
		<td class="all">
			1200
		</td>
		<td class="all">
			1000
		</td>
		<td class="all">
			4000
		</td>
		<td class="all" valign="top">
			20 Ticks
		</td>
		<td class="all" valign="top">
<% if resprogress = 25 then %>
			<%=restime%> Ticks Remaining
<% 
elseif resprogress <> 0 then
	if res25 = 1 then 
%>
			Research Complete
<%
	else
%>
			<font color="#ff0000">Other Research in Progress</font>
<%
	end if 
elseif res25 = 1 then 
%>
			Research Complete
<% elseif res25 = 0 then %>
			<a href="Research_engine.asp?ID=25">Start</a>		
<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>Some more clever tricks used by SEO.</i></small>
		</td>
	</tr>
<% end if %>
</table>			
<!---#include file="includes/footer.asp"--->			
<!---#include file="includes/end_check.asp"--->