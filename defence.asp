<!---#include file="includes\check.asp"--->
<!---#include file="includes\top.asp"--->

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
	<%
	select case Request.QueryString("error")
		case 1
			Response.Redirect("The tick engine is currently running, you can not order new units during this time, please try again in a few seconds.  <BR /><BR />")
	end select
	%>
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
				<%
				Dim rsdefence
					
				set rsdefence = server.CreateObject("ADODB.Recordset")
				strsql = "SELECT NumPilboxes, NumAAGuns, NumBalloons, NumWaterMines, NumTurrets, ConDC, ConRS, ResCS, ResWMP, ResFG FROM tblAccounts WHERE username = '" & strusername & "';"
				
				rsdefence.Open strsql, adocon
				%>
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
						<%
						If rsdefence("ConDC") = 3 Then
						%>
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
								<% =formatnumber(rsdefence("NumPilboxes"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" maxlength="6" name="pillboxes" class="form" value="0" />
							</TD>
						</TR>
						<%
						End If
						If rsdefence("ConRS") = 3 Then
						%>
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
								<% =formatnumber(rsdefence("NumAAGuns"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="aaguns" maxlength="6" class="form" />
							</TD>
						</TR>
						<%
						End If
						If rsdefence("ResCS") = 3 Then
						%>
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
								<% =formatnumber(rsdefence("NumBalloons"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="balloons" maxlength="6" class="form" />
							</TD>
						</TR>
						<%
						End If
						If rsdefence("ResWMP") = 3 Then
						%>
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
								<% =formatnumber(rsdefence("NumWaterMines"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="mines" maxlength="6" class="form" />
							</TD>
						</TR>
						<%
						End If
						If rsdefence("ResFG") = 3 Then
						%>
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
								<% =formatnumber(rsdefence("NumTurrets"), 0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="3" value="0" name="turrets" maxlength="6" class="form" />
							</TD>
						</TR>
						<%
						End If
						%>
						</table>
					<BR />
					<input type="submit" value="Order Defence" class="form">
				</form>
				<%
				rsdefence.Close 
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
				Below is a table of all the units you have ordered.
				<BR />
				<table cellpadding="2" cellspacing="0" border="0">
					<TR>
						<TD class="tdheaderline">
							Unit Name
						</TD>
						<%
						Dim intCounter
						intCounter =1 
						
						for intCounter = 1 to 4
							Response.Write("<TD class=""tdheaderline"" align=""center"" width=""17px"">")
							Response.Write(intCounter)
							Response.Write("</TD>")
						next
						%>
						
					</TR>
					<%
					
					
					strsql = "SELECT timeleft, Pilboxes, AAGuns, Balloons, Mines, Turrets FROM TblPendingUnits WHERE username = '" & strusername & "' ORDER by timeleft ASC;"
					
					rsdefence.open strsql, adocon
					
					Dim strUnitArrival(5,5)
					Dim intProcessTimes(5)
					Dim yCounter
					Dim xCounter
					
					yCounter = 1
					xCounter = 1
					for xcounter = 0 to 5
						for ycounter = 0 to 5
							strUnitArrival(xcounter,ycounter) = "&nbsp;"
						next
					next
					strUnitArrival(0,0) = "Pillboxes"
					strUnitArrival(1,0) = "AA Guns"
					strUnitArrival(2,0) = "Barrage Balloons"
					strUnitArrival(3,0) = "Mines"
					strUnitArrival(4,0) = "Turrets"					
					
					intProcessTimes(0) = 3
					intProcessTimes(1) = 4
					intProcessTimes(2) = 3
					intProcessTimes(3) = 3
					intProcessTimes(4) = 4
					
					for ycounter = 1 to 5
						if not rsdefence.EOF then
								if rsdefence("timeleft") = ycounter then
										if clng(rsdefence("Pilboxes")) <> 0 then
												strunitArrival(0,ycounter) = rsdefence("Pilboxes")
										end if
										if clng(rsdefence("AAGuns")) <> 0 then
												strunitArrival(1,ycounter) = rsdefence("AAGuns")
										end if
										if clng(rsdefence("Balloons")) <> 0 then
												strunitArrival(2,ycounter) = rsdefence("Balloons")
										end if
										if clng(rsdefence("Mines")) <> 0 then
												strunitArrival(3,ycounter) = rsdefence("Mines")
										end if
										if clng(rsdefence("Turrets")) <> 0 then
												strunitArrival(4,ycounter) = rsdefence("Turrets")
										end if
										rsdefence.MoveNext
								end if
						end if
					next
					
					rsdefence.Close 
					set rsdefence = nothing
					%>
					<TR>
					<%
					ycounter = 0
					xcounter = 0
					for xcounter = 0 to 4
						Response.Write("<TR>")
						for ycounter = 0 to 5
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