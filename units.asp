<!---#include file="includes\check.asp"--->
<!---#include file="includes\top.asp"--->

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
	<%
	select case Request.QueryString("error")
		case 1
			Response.write("You have not entered any units to order.<BR /><BR />")
		case 2
			Response.write("You do not have the required resources to place that order.<BR /><BR />")
		case 3
			Response.Write("The requested units have been ordered.<BR /><BR />")
		case 4
			Response.Write("The tick engine is currently processing, you can not order more units during this time, please try again in a few seconds.<BR /><BR />")
			
	end select
	%>
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
				<%
				Dim rsUnits
					
				set rsUnits = server.CreateObject("ADODB.Recordset")
				strsql = "SELECT numinfantry, numcommandos, numjeeps, nummat1, nummat2, numsherman, numspit1, numspit2, numspit5, numhurr, numhalifax, numwellington, numlancaster, numlandingcraft, numlandingship, numfrigates, numbattleship, numcarrier, conJF, ConAA, ConMF, ResBGP, ConSF, ConSPF, ResWC, ResUWR, ConHMP, ConHF, ConWF, ConLF, ResG, ResSIT, ResBLC, ConMP, ResSBS, ResMBS, ResAAS FROM tblaccounts WHERE username = '" & strusername & "';"
				
				rsUnits.Open strsql, adocon
				%>
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
								<% =formatnumber(rsunits("numinfantry"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" maxlength="6" name="infantry" class="form" value="0" />
							</TD>
						</TR>
						<% 
						if rsunits("ConAA") = 3 then
						%>
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
								<% =formatnumber(rsunits("numcommandos"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="comandoes" maxlength="6" class="form" />
							</TD>
						</TR>
						<%
						end if
						if rsunits("conJF") = 3 then
						%>
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
								<% =formatnumber(rsunits("numjeeps"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="jeeps" maxlength="6" class="form" />
							</TD>
						</TR>
						<%
						end if
						if rsunits("ConMF") = 3 then
						%>
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
								<% =formatnumber(rsunits("nummat1"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="matilda1" maxlength="6" class="form" />
							</TD>
						</TR>
						<%
						if rsunits("resBGP") = 3 then
						%>
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
								<% =formatnumber(rsunits("nummat2"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="matilda2" maxlength="6" class="form" />
							</TD>
						</TR>
						<%
						end if
						end if
						if rsunits("ConSF") = 3 then
						%>
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
								<% =formatnumber(rsunits("numsherman"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="sherman" maxlength="4" class="form" />
							</TD>
						</TR>
						<%
						end if
						if rsunits("ConSPF") = 3 then
						%>
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
								<% =formatnumber(rsunits("numspit1"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="spitfire1" maxlength="4" class="form" />
							</TD>
						</TR>
						<%
						if rsunits("ResWC") = 3 then
						%>
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
								<% =formatnumber(rsunits("numspit2"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="spitfire2" maxlength="4" class="form" />
							</TD>
						</TR>
						<% if rsunits("ResUWR") = 3 then %>
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
								<% =formatnumber(rsunits("numspit5"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="spitfire5" maxlength="4" class="form" />
							</TD>
						</TR>
						<%
						end if
						end if
						end if
						if rsunits("ConHMP") = 3 then
						%>
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
								<% =formatnumber(rsunits("numhurr"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="hurricane" maxlength="4" class="form" />
							</TD>
						</TR>
						<%
						end if
						if rsUnits("ConHF") = 3 then
						%>
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
								<% =formatnumber(rsunits("numhalifax"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="halifax" maxlength="4" class="form" />
							</TD>
						</TR>
						<% 
						end if
						if rsunits("ConWF") = 3 then
						%>
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
								<% =formatnumber(rsunits("numwellington"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="wellington" maxlength="4" class="form" />
							</TD>
						</TR>
						<%
						end if
						if (rsunits("conLF") = 3) AND (rsunits("resG") = 3) then
						%>
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
								<% =formatnumber(rsunits("numlancaster"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="lancaster" maxlength="4" class="form" />
							</TD>
						</TR>
						<%
						end if
						if rsunits("resSIT") = 3 then
						%>
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
								<% =formatnumber(rsunits("numlandingcraft"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="landingcraft" maxlength="4" class="form" />
							</TD>
						</TR>
						<%
						end if
						if rsunits("resBLC") = 3 then
						%>
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
								<% =formatnumber(rsunits("numlandingship"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="landingship" maxlength="4" class="form" />
							</TD>
						</TR>
						<% 
						end if
						if (rsunits("ConMP") = 3) AND (rsunits("resSBS") = 3) then
						%>
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
								<% =formatnumber(rsunits("numfrigates"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="frigates" maxlength="4" class="form" />
							</TD>
						</TR>
						<%
						end if 
						if (rsunits("ConMP") = 3) AND (rsunits("resMBS") = 3) then
						%>
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
								<% =formatnumber(rsunits("numbattleship"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="battleship" maxlength="4" class="form" />
							</TD>
						</TR>
						<%
						end if
						if (rsunits("ResMBS") = 3) AND (rsunits("ResAAS") = 3) then
						%>
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
								<% =formatnumber(rsunits("numcarrier"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="carriers" maxlength="4" class="form" />
							</TD>
						</TR>
						<% end if %>
					</table>
					<BR />
					<input type="submit" value="Order Units" class="form">
				</form>
				
				<%
				rsUnits.Close 
				%>
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
						<%
						Dim intCounter
						intCounter =1 
						
						for intCounter = 1 to 24
							Response.Write("<TD class=""tdheaderline"" align=""center"" width=""17px"">")
							Response.Write(intCounter)
							Response.Write("</TD>")
						next
						%>
						
					</TR>
					<%
					
					
					strsql = "SELECT timeleft, infantry, comandoes, jeeps, matilda1, matilda2, sherman, spitfire1, spitfire2, spitfire5, hurricane, halifax, wellington, lancaster, landingcraft, landingship, frigates, battleship, carrier FROM tblpendingunits WHERE username = '" & strusername & "' ORDER by timeleft ASC;"
					
					rsUnits.open strsql, adocon
					
					Dim strUnitArrival(18,25)
					Dim intProcessTimes(18)
					Dim yCounter
					Dim xCounter
					
					yCounter = 1
					xCounter = 1
					for xcounter = 0 to 17
						for ycounter = 0 to 24
							strUnitArrival(xcounter,ycounter) = "&nbsp;"
						next
					next
					strUnitArrival(0,0) = "Infantry"
					strUnitArrival(1,0) = "Commandoes"
					strUnitArrival(2,0) = "Jeeps"
					strUnitArrival(3,0) = "Matilda MkI"
					strUnitArrival(4,0) = "Matilda MkII"
					strUnitArrival(5,0) = "Sherman"
					strUnitArrival(6,0) = "Spitfire MkI"
					strUnitArrival(7,0) = "Spitfire MkII"
					strUnitArrival(8,0) = "Spitfire MkV"
					strUnitArrival(9,0) = "Hurricance MkI"
					strUnitArrival(10,0) = "Halifax Bomber" 
					strUnitArrival(11,0) = "Wellington Bomber"
					strUnitArrival(12,0) = "Lancaster Bomber"
					strUnitArrival(13,0) = "Landing Craft"
					strUnitArrival(14,0) = "Landing Ship" 
					strUnitArrival(15,0) = "Frigates"
					strUnitArrival(16,0) = "Battleships"
					strUnitArrival(17,0) = "Aircraft Carriers"
					
					
					intProcessTimes(0) = 3
					intProcessTimes(1) = 4
					intProcessTimes(2) = 6
					intProcessTimes(3) = 8
					intProcessTimes(4) = 8
					intProcessTimes(5) = 10
					intProcessTimes(6) = 8
 					intProcessTimes(7) = 10
					intProcessTimes(8) = 12
					intProcessTimes(9) = 7
					intProcessTimes(10) = 10
					intProcessTimes(11) = 12
					intProcessTimes(12) = 15
					intProcessTimes(13) = 5
					intProcessTimes(14) = 10
					intProcessTimes(15) = 16
					intProcessTimes(16) = 20
					intProcessTimes(17) = 24
					
					
					for ycounter = 1 to 24
						if not rsUnits.EOF then
								if rsUnits("timeleft") = ycounter then
										if clng(rsunits("infantry")) <> 0 then
												strunitArrival(0,ycounter) = rsunits("infantry")
										end if
										if clng(rsunits("comandoes")) <> 0 then
												strunitArrival(1,ycounter) = rsunits("Comandoes")
										end if
										if clng(rsunits("jeeps")) <> 0 then
												strunitArrival(2,ycounter) = rsunits("Jeeps")
										end if
										if clng(rsunits("matilda1")) <> 0 then
												strunitArrival(3,ycounter) = rsunits("matilda1")
										end if
										if clng(rsunits("matilda2")) <> 0 then
												strunitArrival(4,ycounter) = rsunits("matilda2")
										end if
										if clng(rsunits("sherman")) <> 0 then
												strunitArrival(5,ycounter) = rsunits("sherman")
										end if
										if clng(rsunits("spitfire1")) <> 0 then
												strunitArrival(6,ycounter) = rsunits("spitfire1")
										end if
										if clng(rsunits("spitfire2")) <> 0 then
												strunitArrival(7,ycounter) = rsunits("spitfire2")
										end if
										if clng(rsunits("spitfire5")) <> 0 then
												strunitArrival(8,ycounter) = rsunits("spitfire5")
										end if
										if clng(rsunits("hurricane")) <> 0 then
												strunitArrival(9,ycounter) = rsunits("hurricane")
										end if
										if clng(rsunits("halifax")) <> 0 then
												strunitArrival(10,ycounter) = rsunits("halifax")
										end if
										if clng(rsunits("wellington")) <> 0 then
												strunitArrival(11,ycounter) = rsunits("wellington")
										end if
										if clng(rsunits("lancaster")) <> 0 then
												strunitArrival(12,ycounter) = rsunits("lancaster")
										end if
										if clng(rsunits("landingcraft")) <> 0 then
												strunitArrival(13,ycounter) = rsunits("landingcraft")
										end if
										if clng(rsunits("landingship")) <> 0 then
												strunitArrival(14,ycounter) = rsunits("landingship")
										end if
										if clng(rsunits("frigates")) <> 0 then
												strunitArrival(15,ycounter) = rsunits("frigates")
										end if
										if clng(rsunits("battleship")) <> 0 then
												strunitArrival(16,ycounter) = rsunits("battleship")
										end if
										if clng(rsunits("carrier")) <> 0 then
												strunitArrival(17,ycounter) = rsunits("carrier")
										end if
										rsUnits.MoveNext
								end if
						end if
					next
					
					rsUnits.Close 
					set rsunits = nothing
					%>
					<TR>
					<%
					ycounter = 0
					xcounter = 0
					for xcounter = 0 to 17 
						Response.Write("<TR>")
						for ycounter = 0 to 24
							if intProcessTimes(xcounter) >= ycounter then
									Response.Write("<TD class=form align=center>")
									Response.Write(strunitarrival(xcounter,ycounter))
									Response.Write("</TD>")
								else
							end if

						next
						Response.Write("</TR>")
					next
					%>
					</TR>
				</table>
			</TD>
		</TR>
	</table>
	<BR />
</center>
<!---#include file="includes\footer.asp"--->
<!---#include file="includes\end_check.asp"--->