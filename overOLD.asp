
<!---#include file="includes\check.asp"--->
<!---#include file="includes\top.asp"--->
<center>
	<img SRC="images/overview.gif" border="0" alt="Overview" />

<%
Dim rsIncomming		'holds the unit movements

strSQL= "SELECT ToCountry, tocontinent FROM TblMovements WHERE status = 1 AND toContinent = " & lngContinentNo

set rsIncomming = server.CreateObject("ADODB.Recordset")

rsIncomming.Open strsql, adocon

if rsIncomming.EOF then
else
%>
<BR />
<font class="norm" color="#FF0000">
	<B>
		<U>
			YOUR CONTINENT HAS INCOMING ATTACKERS!
		</u>
	</B>
</FONT>
<BR />
<%
end if

rsIncomming.Close
%>
<BR />
<BR />
<table cellpadding="2" cellpadding="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
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
<A HREF="admin.asp">Click here to use the admin console</A>
<br />

<%
end if
%>
<BR />
<table cellpadding="2" cellpadding="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
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
<table cellpadding="2" cellpadding="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline" align="center">
			Incomming Units
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
								strincommingtxt = "#00FFFF"
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
							
							response.Write("Incomming Units:"& dblunitstotal)
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

strsql = "SELECT researchticks, conticks, numPoW, NumInfantry, NumCommandos, NumJeeps, NumMat1, NumMat2, NumSherman, NumSpit1, NumSpit2, NumSpit5, NumHurr, NumLancaster, NumWellington, NumHalifax, NumLandingCraft, NumLandingShip, NumFrigates, NumBattleShip, NumCarrier, NumCargoShip, NumAAGuns, NumPilboxes, NumTurrets, NumBalloons, NumWaterMines, NumMP, NumSpies, NumExplorers, ProtectionTime, ResSF, ResCE, ResVGM, resBGP, ResMPV, ResMWP, ResMPC, ResFPC, ResWC, ResRP, ResUWR, ResG, ResEP, ResLB, ResSIT, ResBLC, ResSBS, ResMBS, ResAAS, ResAES, ResCS, ResFG, ResWMP, ResE, ResD, ResAC, ConAA, ConJF, ConMF, ConSF, ConSPF, ConHMP, ConLf, ConHF, ConWF, ConBF, ConRF, ConMP, ConDC, ConRS, ConSC, ConSRC, NumFarms, NumMines, NumRefineries, NumMills FROM tblaccounts WHERE username = '" & strusername & "';"

rsIncomming.Open strsql, adocon
%>	
<table cellpadding="2" cellpadding="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
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
									Comandoes:
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
									Sherman Tank:
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
									Halifax Bomber:
								</TD>
								<TD class="norm">
									<% =rsincomming("numhalifax") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Wellington Bomber:
								</TD>
								<TD class="norm">
									<% =rsincomming("numwellington") %>
								</TD>
							</TR>
							<TR>
								<TD align="right" class="norm">
									Lancaster Bomber:
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
									Landing Ship:
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
									Battle Ships:
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
							<TR>
								<TD align="right" class="norm">
									Cargo Ships:
								</TD>
								<TD class="norm">
									<% =rsincomming("numcargoship") %>
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
							<TR>
								<TD align="right" class="norm">
									PoWs:
								</TD>
								<TD class="norm">
									<% =rsincomming("numpow") %>
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
<table cellpadding="2" cellpadding="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
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
			Your currently researching
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
			and will finish in <% =rsincomming("researchticks") %> time.  	
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
			select case rsincomming("CurCon")
				case "AA"
					Response.Write("Advance Acadamy")
				case "JF"
					Response.Write("Jeep Factory")
				case "MF"
					Response.Write("Matilda Factory")
				case "SF"
					Response.Write("Sherman Factory")
				case "SPF"
					Response.Write("Spitfire Factory")
				case "HMP"
					Response.Write("Hawlker Manufacture Plant")
				case "LF"
					Response.Write("Lancaster Factory")
				case "HF"
					Response.Write("Halifax Factory")
				case "WF"
					Response.Write("Wellington Factory")
				case "BF"
					Response.Write("Bomb Factory")
				case "RF"
					Response.Write("Rocket Factory")
				case "MP"
					Response.Write("Main Port")
				case "DC"
					Response.Write("Defence Centre")
				case "RS"
					Response.Write("Radar Station")
				case "SC"
					Response.Write("Spy Training Centre")
				case "SRC"
					Response.Write("Spy Resorce Centre")
			end select
			%>
			and will finish in <% =rsincomming("conticks") %> time.
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
