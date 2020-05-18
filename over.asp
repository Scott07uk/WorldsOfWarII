
<!---#include file="includes\check.asp"--->
<!---#include file="includes\top.asp"--->
<!---#include file="includes\stdFunctions.asp"--->

<!---SPELL CHECKED--->
<SCRIPT language="javascript">
			function check(formobj){
				if ((document.frmform.option1.value == document.frmform.option2.value) || (document.frmform.option1.value == document.frmform.option3.value) || (document.frmform.option2.value == document.frmform.option1.value) || (document.frmform.option2.value == document.frmform.option3.value) || (document.frmform.option3.value == document.frmform.option1.value) || (document.frmform.option3.value == document.frmform.option2.value) || (document.frmform.option1.value =="") || (document.frmform.option2.value =="") || (document.frmform.option2.value =="")) {
					alert ("Cannot choose same option!");
				} else {
					formobj.submit();
				}
			}
			function option1check(selected){
				if ((selected == document.frmform.option2.value) || (selected == document.frmform.option3.value)) {
						document.frmform.option1.value = "---";
				}
			}
			function option2check(selected){
				if (selected == document.frmform.option1.value || selected == document.frmform.option3.value) {
					document.frmform.option2.value = "---";
				}
			}
			function option3check(selected){
				if (selected == document.frmform.option1.value || selected == document.frmform.option2.value) {
					document.frmform.option3.value = "---";
				}
			}
</script>

<center>
<table width="95%" cellpadding="0" border="0" cellspacing="0" style="border-top:none; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD align="center" class="norm">
			<A HREF="javascript:openHelp('1');">Help With This Page</A>
		</TD>
	</TR>
</table>
</center>
<%
Dim rsIncomming		'holds the unit movements

strSQL= "SELECT ToCountry, tocontinent FROM TblMovements WHERE status = 1 AND toContinent = " & lngContinentNo

set rsIncomming = server.CreateObject("ADODB.Recordset")

rsIncomming.Open strsql, adocon

if rsIncomming.EOF then
else
%>
<BR />
<font color="#FF0000">
	<B>
		<U>
			<center>YOUR CONTINENT HAS INCOMING ATTACKERS</center>
		</u>
	</B>
</FONT>
<%
end if

rsIncomming.Close
%>
<BR />
<table align="center" cellpadding="2" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline" align="center">
			Message from Creators
		</TD>
	</TR>
	<TR>
		<TD align="left" class="norm">
			<%
			Dim rsConfig
			
			strSQL = "SELECT tblconfig.CreatorsMessage FROM tblconfig;"
			
			'create an object to load the message from commanders
			set rsConfig = server.CreateObject("ADODB.Recordset")
			'open the db
			rsConfig.open strSQL, adocon
			if rsConfig.EOF then
			%>
				Welcome to Worlds of War II
			<%
			else
			'print out the message from creators
			Response.Write(rsconfig("creatorsmessage"))
			end if
			
			rsconfig.close
			set rsconfig = nothing
			%>
		</TD>
	</TR>
</table>
<BR />
<%
if blnadmin = true then
%>
<center><A HREF="admin.asp">ADMIN CONSOLE</A></center>
<br />

<%
end if
%>
<BR />
<table align="center" cellpadding="2" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline" align="center">
			Worlds of War II Game Ranking
		</TD>
	</TR>
	<TR>
		<TD align="left" class="norm" align="center">
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
<form name="frmform" action="voteprocess.asp" method="post">
	<table align="center" cellpadding="2" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
		<TR>
			<TD class="tdheaderline" align="center">
				Redevelopment ideas!
			</TD>
		</TR>
		<TR>
			<TD align="center" class="norm">
<%
Dim rsVotes
set rsuser = server.CreateObject("ADODB.recordset")
set rsvotes = server.CreateObject("ADODB.recordset")

strsql = "SELECT tbluservotes.username FROM tbluservotes WHERE username = '" & strUsername & "';"

rsuser.Open strSql, adocon

if not rsuser.EOF then

	strsql = "SELECT TOP 3 tblideavotes.ideaNo, name  FROM tblideavotes ORDER BY votescore DESC;"
					
	rsvotes.Open strSql, adocon

	if not rsvotes.EOF then
%>
	The top three ideas so far are...
<%
		for i = 1 to 3
%>
			<br>
			<%=i%>. <%=rsvotes("name")%>
<%			rsvotes.movenext
		next

	end if
	
	rsvotes.close
	set rsvotes = nothing
else

	strsql = "SELECT tblideavotes.ideaNo, name  FROM tblideavotes;"
					
	rsvotes.Open strSql, adocon
	Dim strOptions
	stroptions = "<OPTION></OPTION>"
	if not rsvotes.EOF then
	%>
					Please select your first, second and third favorite ideas!
					<select style="width: 100px;" class="form" name="option1" id="option1" onchange="javascript:option1check(this.value);">
	<%
		do while not rsvotes.EOF
			strOptions = strOptions & "<OPTION VALUE=""" & RSVOTES("IDEANO") & """>" & RSVOTES("NAME") & "</OPTION>"
			RSVOTES.MOVENEXT
		LOOP
	%>
						<% =strOptions %>
					</select>

					<select style="width: 100px;" name="option2" class="form" id="option2" onchange="javascript:option2check(this.value);">
						<% =strOptions %>
					</select>

					<select style="width: 100px;" name="option3" class="form" id="option3" onchange="javascript:option3check(this.value);">
						<% =strOptions %>
					</select>
	<%
	end if
	rsvotes.Close
	set rsvotes = nothing
%>
				</select>
				<input type="button" value="Vote!" name="vote" class="form" onclick="javascript: check(this.form);">
				<br />
				<br />
				Click <a href="http://www.fhforums.co.uk/forum_posts.asp?TID=903&PN=1" target="_blank">here</a> to view details about all the ideas.
<%
end if
rsuser.close
%>
			</TD>
		</TR>
	</table>
</form>
<BR />
<BR />
<table align="center" cellpadding="2" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline" align="center">
			Message from Commanders
		</TD>
	</TR>
	<TR>
		<TD align="left" class="norm">
			<%
			Dim rsMOTD
			
			strSQL = "SELECT tblcontinents.message, continentno, name, banner FROM tblcontinents WHERE continentno = " & lngcontinentno
			
			'create an object to load the message from commanders
			set rsMOTD = server.CreateObject("ADODB.Recordset")
			
			'set the update properties to allow the db to be updated
			rsMOTD.LockType = 3
			rsMOTD.CursorType = 2
			
			'open the db
			rsMOTD.open strSQL, adocon
			if rsMOTD.EOF then
				rsMOTD.AddNew
				rsMOTD.Fields("ContinentNo") = lngcontinentno
				rsMOTD.Fields("name") = "Worlds of War II"
				rsMOTD.Fields("banner") = "images/none.gif"
				rsMOTD.update
				
			%>
				No message has been set.
			<%
			else
			'print out the message from creators
			Response.Write(rsMOTD("message"))
			end if
			
			rsMOTD.close
			set rsMOTD = nothing
			%>
		</TD>
	</TR>
</table>
<BR />
<BR />
<%
strsql = "SELECT ETA, infantry, comandoes, jeeps, matilda1, matilda2, sherman, spitfire1, spitfire2, spitfire5, hurricane, lancaster, wellington, halifax, frigates, battleships, carrier, status, fromcountry, fromcontinent FROM tblmovements WHERE tocontinent = " & lngcontinentno & " AND tocountry = " & lngcountryno & " ORDER BY ETA ASC;"

rsIncomming.Open strsql, adocon

if not rsIncomming.EOF then
%>
<table align="center" cellpadding="2" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline" align="center">
			Incoming Units
		</TD>
	</TR>
	<TR>
		<TD align="left" class="norm">
			<table cellpadding="2" cellspacing="0" border="0" width="100%">
				<%
				Dim strincommingtxt		'varible to hold color of text
				Dim strMissionType		'varible to hold the type of mission
				do while not rsIncomming.EOF
				if rsincomming("status") = 1 then
						strincommingtxt = "#FF0000"
						strMissionType = "Attack"
					else
						if rsincomming("status") = 2 then
								strincommingtxt = "#00FF00"
								strMissionType = "Defend"
							else
								strincommingtxt = "#FFFFFF"
								strMissionType = "Return"
						end if
				end if
				%>
				<TR>
					<TD class="norm">
						<font color="<% =strincommingtxt %>">
							<% =rsincomming("fromcontinent") %>:<% =rsincomming("fromcountry") %>
						</font>
					</td>
					<td class="norm">
						<font color="<% =strincommingtxt %>">
							ETA: <% =rsincomming("eta") %>
						</font>
					</td>
					<td class="norm">
						<font color="<% =strincommingtxt %>">
							Mission: <% =strMissionType %>
						</font>
					</td>
					<td class="norm">
						<font color="<% =strincommingtxt %>">
							<%
							Dim dblunitstotal		'varible to hold the total incomming units
							dblunitstotal = cdbl(rsincomming("infantry"))+cdbl(rsincomming("comandoes"))+cdbl(rsincomming("jeeps"))+cdbl(rsincomming("matilda1"))+cdbl(rsincomming("matilda2"))+cdbl(rsincomming("sherman"))+cdbl(rsincomming("spitfire1"))+cdbl(rsincomming("spitfire2"))+cdbl(rsincomming("spitfire5"))+cdbl(rsincomming("hurricane"))+cdbl(rsincomming("lancaster"))+cdbl(rsincomming("wellington"))+cdbl(rsincomming("halifax"))+cdbl(rsincomming("frigates"))+cdbl(rsincomming("battleships"))+cdbl(rsincomming("carrier"))
							
							response.Write("Incoming Units:"& dblunitstotal)
							%>
						</font>
					</td>
				</tr>
				<%
				rsincomming.movenext
				loop
				%>
			</table>
		</TD>
	</TR>
</table>
<BR />
<BR />
<%
end if

rsincomming.close

strsql = "SELECT numinfantry, numcommandos, numjeeps, nummat1, nummat2, numsherman, numspit1, numspit2, numspit5, numhurr, numhalifax, numwellington, numlancaster, numlandingship, numlandingcraft, numfrigates, numbattleship, numcarrier, numpilboxes, numaaguns, numturrets, numballoons, numwatermines, numexplorers, nummp, numspies, numfarms, nummines, nummills, numrefineries, curresearch, curcon, researchticks, conticks FROM tblaccounts WHERE username = '" & strusername & "';"

rsIncomming.Open strsql, adocon
%>	
<table align="center" cellpadding="2" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline" align="center">
			Units
		</TD>
	</TR>
	<TR>
		<TD align="left" class="norm">
			<table cellpadding="2" width="100%" border="0" cellspacing="0">
				<TR>
					<TD align="center" valign="top" class="norm" width="50%">
						<B>
							<U>
								Mobile</U>
						</B>
						<BR />
						<table cellpadding="2" border="0" cellspacing="0">
							<TR>
								<TD align="right" class="norm" width="70%">
									Infantry:
								</TD>
								<TD class="norm" width="30%">
									<% =rsincomming("numinfantry") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Commandos:
								</TD>
								<TD class="norm">
									<% =rsincomming("numcommandos") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Jeeps:
								</TD>
								<TD class="norm">
									<% =rsincomming("numjeeps") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Matilda MkI:
								</TD>
								<TD class="norm">
									<% =rsincomming("nummat1") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Matilda MkII:
								</TD>
								<TD class="norm">
									<% =rsincomming("nummat2") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Sherman Tanks:
								</TD>
								<TD class="norm">
									<% =rsincomming("numsherman") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Supermarine Spitfire MkI:
								</TD>
								<TD class="norm">
									<% =rsincomming("numspit1") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Supermarine Spitfire MkII:
								</TD>
								<TD class="norm">
									<% =rsincomming("numspit2") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Supermarine Spitfire MkV:
								</TD>
								<TD class="norm">
									<% =rsincomming("numspit5") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Hawlker Hurricane MkI:
								</TD>
								<TD class="norm">
									<% =rsincomming("numhurr") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Halifax Bombers:
								</TD>
								<TD class="norm">
									<% =rsincomming("numhalifax") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Wellington Bombers:
								</TD>
								<TD class="norm">
									<% =rsincomming("numwellington") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Lancaster Bombers:
								</TD>
								<TD class="norm">
									<% =rsincomming("numlancaster") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Landing Craft:
								</TD>
								<TD class="norm">
									<% =rsincomming("numlandingcraft") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Landing Ships:
								</TD>
								<TD class="norm">
									<% =rsincomming("numlandingship") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Frigates:
								</TD>
								<TD class="norm">
									<% =rsincomming("numfrigates") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Battleships:
								</TD>
								<TD class="norm">
									<% =rsincomming("numbattleship") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Aircraft Carriers:
								</TD>
								<TD class="norm">
									<% =rsincomming("numcarrier") %>
								</TD>
							</TR>
						</table>
					</TD>
					<TD align="center" valign="top" width="50%">
						<B>
							<U>
								Fixed Units</U>
						</B>
						<BR />
						<table width="100%" cellpadding="2" border="0" cellspacing="0">
							<TR>
								<TD align="right" class="norm" width="55%">
									Pillboxes:
								</TD>
								<TD class="norm" width="45%">
									<% =rsincomming("numpilboxes") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Turret Guns:
								</TD>
								<TD class="norm">
									<% =rsincomming("numturrets") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									AA Guns:
								</TD>
								<TD class="norm">
									<% =rsincomming("numaaguns") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Barrage Balloons:
								</TD>
								<TD class="norm">
									<% =rsincomming("numballoons") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Water Mines:
								</TD>
								<TD class="norm">
									<% =rsincomming("numwatermines") %>
								</TD>
							</TR>
						</Table>
						<BR />
						<BR />
						<B>
							<U>
								Other Units</U>
						</B>
						<BR />
						<table cellpadding="2" width="100%" border="0" cellspacing="0">
							<TR>
								<TD align="right" class="norm" width="55%">
									Military Police:
								</TD>
								<TD class="norm" width="45%">
									<% =rsincomming("nummp") %>
								</TD>
							</TR>
							<TR>
								<td align="right" class="norm">
									Spies:
								</TD>
								<TD class="norm">
									<% =rsincomming("numspies") %>
								</TD>
							</TR>
							<TR>
								<td align="right" class="norm">
									Explorers:
								</TD>
								<TD class="norm">
									<% =rsincomming("numexplorers") %>
								</TD>
							</TR>
						</table>
						<BR />
						<BR />
						<B>
							<U>
								Production Units</U>
						</B>
						<BR />
						<table cellpadding="2" border="0" width="100%" cellspacing="0">
							<TR>
								<TD align="right" class="norm" width="55%">
									Farms:
								</TD>
								<TD class="norm" width="45%">
									<% =rsincomming("NumFarms") %>
								</TD>
							</TR>
							<TR>
								<td align="right" class="norm">
									Mines:
								</TD>
								<TD class="norm">
									<% =rsincomming("nummines") %>
								</TD>
							</TR>
							<TR>
								<td align="right" class="norm">
									Refineries:
								</TD>
								<TD class="norm">
									<% =rsincomming("numrefineries") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Mills:
								</TD>
								<TD class="norm">
									<% =rsincomming("nummills") %>
								</TD>
							</TR>
						</table>
					</TD>
				</TR>
			</table>
		</TD>
	</TR>
</table>
<BR />
<BR />
<table align="center" cellpadding="2" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline" align="center">
			Research and Construction
		</TD>
	</TR>
	<TR>
		<TD align="left" class="norm">
			<%
			if rsincomming("researchticks") = 0 then
			%>
			Your research teams are standing by for your next order.  
			<%
			else
			%>
			You are currently researching
			<%
			select case rsincomming("CurResearch")
				case "CE"
					Response.Write("Combustion Engines for Millitary Use")
				case "SF"
					Response.Write("Special Forces")
				case "VGM"
					Response.Write("Vehicle Gun Mounting")
				case "BGP"
					Response.Write("Balancing large Gun Power")
				case "MPV"
					Response.Write("Mass Production of Military Vehicles")
				case "MWP"
					Response.Write("Mono-Wing Planes")
				case "MPC"
					Response.Write("Metal Plane Construction")
				case "FPC"
					Response.Write("Fabric Plane Covers")
				case "WC"
					Response.Write("Wing Cannons")
				case "RP"
					Response.Write("Rocket Purpolsion")
				case "UWR"
					Response.Write("Under Wing Rockets")
				case "G"
					Response.Write("G")
				case "EP"
					Response.Write("Efficiant Payloads")
				case "LB"
					Response.Write("Larger Bombs")
				case "SIT"
					Response.Write("Sea Invasion Tatics")
				case "BLC"
					Response.Write("Bigger Landing Craft")
				case "SBS"
					Response.Write("Small Scale Morden Battleships")
				case "MBS"
					Response.Write("Morden Full Size Battleships")
				case "AAS"
					Response.Write("Aviation at Sea")
				case "AES"
					Response.Write("Auto Exploding Shells")
				case "CS"
					Response.Write("Chemical Use")
				case "FG"
					Response.Write("Fixed Guns")
				case "WMP"
					Response.Write("Water Mine Patturns")
				case "E"
					Response.Write("Encryption")
				Case "D"
					Response.Write("Deception")
				Case "AC"
					Response.Write("Advanced Communications")
			end select
			%>
			and will finish in <% =rsincomming("researchticks") %> ticks.  	
			<%
			end if
			%>
			<BR />
			<%
			if rsincomming("conTicks") = 0 then
			%>
			Your construction teams are awaiting your orders.  
			<%
			else
			%>
			Your construction teams are currently building
			<%
			Response.Write(rtnConName(rsincomming("curcon")))
			%>
			and will finish in <% =rsincomming("conticks") %> ticks.
			<%
			end if
			%>  	
		</TD>
	</TR>
</table>
<BR />
<BR />
<%
rsIncomming.Close

set rsincomming = nothing
%>		
<!---#include file="includes\footer.asp"--->
<!---#include file="includes\end_check.asp"--->
