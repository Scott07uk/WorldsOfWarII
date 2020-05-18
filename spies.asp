<!---#include file="includes\check.asp"--->
<!---#include file="includes\top.asp"--->
<center>
<table width="95%" cellpadding="0" border="0" cellspacing="0" style="border-top:none; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD align="center" class="norm">
			<A HREF="javascript:openHelp('15');">Help With This Page</A>
		</TD>
	</TR>
</table>
</center>
<%
'SPELL CHECKED

Dim rsSpies

strsql = "SELECT accuracy, reportid, countryid, continentid, tickreport FROM tblintelegencereports WHERE username = '" & strusername & "';"

set rsSpies = server.CreateObject("ADODB.Recordset")

rsSpies.Open strsql, adocon

%>
<center>
	<BR />
	<BR />
	<%
	select case Request.QueryString("error")
		case 1
			Response.Write("You did not enter any units to order.<BR /><BR />")
		case 2
			Response.write("You do not have enough food to order that many units.<BR /><BR />")
		case 3
			Response.Write("You do not have enough wood to order that many units.<BR /><BR />")
		case 4
			Response.Write("You do not have enough iron to order that many units.<BR /><BR />")
		case 5
			Response.Write("You do not have enough money to order that many units.<BR /><BR />")
		case 6
			Response.Write("The units you requested have been enrolled for training.<BR /><BR />")
		case 7
			Response.Write("You did not entered a valid number of units to send.<BR /><BR />")
		case 8
			Response.Write("You did not enter a valid target nation.<BR /><BR />")			
		case 9
			Response.Write("The target country is not attacking you. You cannot spy on the forces that they are sending to you.<BR /><BR />")
		case 10
			Response.Write("Unfortunately our spies failed in their operation. A more detailed report can be found in news.<BR /><BR />")
		case 11
			Response.Write("The mission was successful.<BR /><BR />")
		case 12
			Response.Write("The intelligence report has been deleted.<BR /><BR />")
		case 13
			Response.Write("You do not have enough spies to launch that mission.<BR /><Br />")
		case 14
			Response.Write("The target you selected has no refineries.<Br /><BR />")
		case 15
			Response.Write("The target you selected has no bombing aircraft.<BR /><BR />")
		case 16
			Response.Write("The tick engine is running, you can not order more spies during this period, please try again in a few seconds.  <BR /><BR />")
			
	end select
	%>
	<table cellpadding="2" cellspacing="0" border="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
		<TR>
			<TD class="tdheaderline" align="center">
				Saved Intelligence Reports
			</TD>
		</TR>
		<TR>
			<TD class="norm" align="center">
				Here are the intelligence reports that have been created from espionage missions.
				<BR />
				<%
				if rsSpies.EOF then
				%>
				You have no reports to show at the moment.  
				<%
				else
				%>
				<table cellpadding="2" border="0" width="100%" cellspacing="0">
					<TR>
						<TD class="tdheaderline">
							Location
						</TD>
						<TD class="tdheaderline">
							Accuracy
						</TD>
						<TD class="tdheaderline">
							Tick Time
						</TD>
						<TD class="tdheaderline">
							View?
						</TD>
						<TD class="tdheaderline">
							Delete?
						</TD>
					</TR>
					<%
					do while not rsSpies.EOF 
					%>
					<TR>
						<TD class="norm">
							<% =rsspies("continentid") %>:<% =rsspies("countryid") %>
						</TD>
						<TD class="norm">
							<% =formatnumber(rsspies("accuracy"),1) %>%
						</TD>
						<TD class="norm">
							<% =formatnumber(rsspies("tickreport"),0) %>
						</TD>
						<TD class="norm">
							<A HREF="view_spy_report.asp?id=<% =rsspies("reportid") %>">View</A>
						</TD>
						<TD class="norm">
							<A HREF="delete_spy_report.asp?id=<% =rsspies("reportid") %>" onclick="return confirm('Are you sure you want to delete this report?');">Delete</A>
						</TD>
					</TR>
					<%
					rsSpies.MoveNext 
					loop
					%>
				</table>
				<%
				end if
				%>
			</TD>
		</TR>
	</table>
	<BR />
	<BR />
	<%
	rsSpies.Close 
	
	strsql = "SELECT numspies, nummp, conMP, conAA, resSIT, ConMP, ResRP, ResCS, resE FROM tblaccounts WHERE username = '" & strusername & "';"
	
	rsSpies.Open strsql, adocon
	%>
	<table width="95%" cellpadding="2" cellspacing="0" border="0" style="border-top:solid; border-bottom:solid; border-right:solid; border-left:solid; border-color:#FFFFFF; border-width:1;">
		<TR>
			<TD class="tdheaderline" align="center">
				Train new units
			</TD>
		</TR>
		<TR>
			<TD class="norm" align="center">
				<form action="order_spy.asp" method="post" name="spy" onsubmit="">
					<table cellpadding="5" border="0" width="100%" cellspacing="0">
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
						<TR title="Military police protect you from spy attacks.">
							<TD class="norm">
								Military Police
							</TD>
							<TD class="norm">
								60
							</TD>
							<TD class="norm">
								0
							</TD>
							<TD class="norm">
								5
							</TD>
							<TD class="norm">
								150
							</TD>
							<TD class="norm">
								6 Ticks
							</TD>
							<TD class="norm">
								<% =formatnumber(rsspies("nummp"),0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="2" name="mp" value="0" class="form" maxlength="3" />
							</TD>
						</TR>
						<%
						if (cint(rsspies("conAA")) = 3) AND (cint(rsspies("ressIT")) = 3) AND (cint(rsspies("conMP")) = 3) AND (cint(rsspies("resRP")) = 3) AND (cint(rsspies("ResCS")) = 3) and (cint(rsspies("rese")) = 3) then
						%>
						<TR title="Spies allow you to gather intelegance on your enemies.">
							<TD class="norm">
								Spies
							</TD>
							<TD class="norm">
								100
							</TD>
							<TD class="norm">
								1
							</TD>
							<TD class="norm">
								5
							</TD>
							<TD class="norm">
								400
							</TD>
							<TD class="norm">
								12 Ticks
							</TD>
							<TD class="norm">
								<% =formatnumber(rsspies("numspies"),0) %>
							</TD>
							<TD class="norm">
								<input type="text" size="2" name="spies" value="0" class="form" maxlength="3" />
							</TD>
						</TR>
						<%
						end if
						%>
					</table>
					<BR />
					<input type="submit" value="Order" class="form">
				</form>
			</TD>
		</TR>
	</table>
	<BR />
	<BR />
	<table cellpadding="2" border="0" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#ffffff; border-width:1;">
		<TR>
			<TD class="tdheaderline" align="center">
				New Operation
			</TD>
		</TR>
		<TR>
			<TD class="norm" align="center">
				Here you can make a new espionage operation against a target.
				<BR />
				<table cellpadding="2" border="0" width="100%" cellspacing="0">
					<TR>
						<TD class="tdheaderline">
							Name
						</TD>
						<TD class="tdheaderline">
							Description
						</TD>
						<TD class="tdheaderline" width="83">
							Target
						</TD>
						<TD class="tdheaderline">
							Send Spies
						</TD>
						<TD class="tdheaderline">
							Send
						</TD>
					</TR>
					<TR>
						<form action="launch_spies.asp" method="post">
							<input type="hidden" name="type" value="1" />
							
							<TD class="norm">
								Incoming Spy
							</TD>
							<TD class="norm">
								<font size="1">
									Send spies to analyse the orders for an attack on you originating from the target.
								</font>
							</TD>
							<TD class="norm">
								<input type="text" size="1" maxlength="4" name="continent" class="form" />:<input type="text" size="1" maxlength="2" name="country" class="form" />
							</TD>
							<TD class="norm">
								
								<input type="text" size="4" name="spies" class="form" maxlength="6" />
							</TD>
							<TD class="norm">
								<input type="submit" value="Send" class="form" />
							</TD>
						</form>
					</TR>
					<TR>
						<form action="launch_spies.asp" method="post">
							<input type="hidden" name="type" value="2" />
							<TD class="norm">
								Base Spy
							</TD>
							<TD class="norm">
								<font size="1">
									Send spies to analyse what units the target currently has at home.
								</font>
							</TD>
							<TD class="norm">
								<input type="text" size="1" maxlength="4" name="continent" class="form" />:<input type="text" size="1" maxlength="2" name="country" class="form" />
							</TD>
							<TD class="norm">
								
								<input type="text" size="4" name="spies" class="form" maxlength="6" />
							</TD>
							<TD class="norm">
								<input type="submit" value="Send" class="form" />
							</TD>
						</form>
					</TR>
					<TR>
						<form action="launch_spies.asp" method="post">
							<input type="hidden" name="type" value="3" />
							<TD class="norm">
								Surging Power
							</TD>
							<TD class="norm">
								<font size="1">
									Send spies to short circuit the electrical equipment in a refinery, causing it to explode in the target country.
								</font>
							</TD>
							<TD class="norm">
								<input type="text" size="1" maxlength="4" name="continent" class="form" />:<input type="text" size="1" maxlength="2" name="country" class="form" />
							</TD>
							<TD class="norm">
								
								<input type="text" size="4" name="spies" class="form" maxlength="6" />
							</TD>
							<TD class="norm">
								<input type="submit" value="Send" class="form" />
							</TD>
						</form>
					</TR>
					<TR>
						<form action="launch_spies.asp" method="post">
							<input type="hidden" name="type" value="4" />
							<TD class="norm">
								Down With Bombing
							</TD>
							<TD class="norm">
								<font size="1">
									Send spies to remove key parts from bomber engines, thus rendering them useless in the target country.
								</font>
							</TD>
							<TD class="norm">
								<input type="text" size="1" maxlength="4" name="continent" class="form" />:<input type="text" size="1" maxlength="2" name="country" class="form" />
							</TD>
							<TD class="norm">
								
								<input type="text" size="4" name="spies" class="form" maxlength="6" />
							</TD>
							<TD class="norm">
								<input type="submit" value="Send" class="form" />
							</TD>
						</form>
					</TR>
				</table>
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
						
						for intCounter = 1 to 12
							Response.Write("<TD class=""tdheaderline"" align=""center"" width=""17px"">")
							Response.Write(intCounter)
							Response.Write("</TD>")
						next
						%>
						
					</TR>
					<%
					rsSpies.Close 
					
					strsql = "SELECT timeleft, spies, mp FROM tblpendingunits WHERE username = '" & strusername & "' ORDER by timeleft ASC;"
					
					rsspies.open strsql, adocon
					
					Dim strUnitArrival(2,13)
					Dim intProcessTimes(2)
					Dim yCounter
					Dim xCounter
					
					yCounter = 1
					xCounter = 1
					for xcounter = 0 to 1
						for ycounter = 0 to 12
							strUnitArrival(xcounter,ycounter) = "&nbsp;"
						next
					next
					strUnitArrival(0,0) = "MP's"
					strUnitArrival(1,0) = "Spies"
					
					
					intProcessTimes(0) = 6
					intProcessTimes(1) = 12
					
					for ycounter = 1 to 12
						if not rsspies.EOF then
								if rsspies("timeleft") = ycounter then
										if clng(rsspies("mp")) <> 0 then
												strunitArrival(0,ycounter) = rsspies("mp")
										end if
										if clng(rsspies("spies")) <> 0 then
												strunitArrival(1,ycounter) = rsspies("spies")
										end if
									
										
										rsspies.MoveNext
								end if
						end if
					next
					%>
					<TR>
					<%
					ycounter = 0
					xcounter = 0
					for xcounter = 0 to 1 
						Response.Write("<TR>")
						for ycounter = 0 to 12
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
<%
rsSpies.Close 

set rsspies = nothing
%>
</center>
					
<!---#include file="includes\footer.asp"--->
<!---#include file="includes\end_check.asp"--->