<%
'SPELL CHECKED

Dim rsspies
Dim lngReportid

'get the user input
lngreportid = Request.QueryString("id")

if isnumeric(lngreportid) = false then
		Response.Redirect("spies.asp")
	else
%>
<!---#include file="includes\check.asp"--->
<!---#include file="includes\top.asp"--->
<BR />
<BR />
<%
set rsspies = server.CreateObject("ADODB.Recordset")

strsql = "SELECT username, accuracy, infantry, comandoes, jeeps, matilda1, matilda2, sherman, spitfire1, spitfire2, spitfire5, hurricane, halifax, wellington, lancaster, landingcraft, landingship, frigates, battleships, carrier, aaguns, pilboxes, turrets, balloons, mines, spies, mp, countryid, continentid, type, tickreport FROM TblIntelegenceReports WHERE reportid = " & lngreportid & " AND username = '" & strusername & "';"

rsspies.Open strsql, adocon

if rsspies.EOF then
%>
<center>
	The report you wanted to view does not exist.
</center>
<%
else
%>
<center>
	<table cellpadding="2" border="0" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
		<TR>
			<TD class="tdheaderline" align="center">
				Intelligence Report On <% =rsspies("continentid") %>:<% =rsspies("countryid") %> Tick <% =formatnumber(rsspies("tickreport"),0) %>.
				<BR />
				Accuracy: <% =formatnumber(rsspies("accuracy"),1) %>%
			</TD>
		</TR>
		<TR>
			<TD class="norm" align="center">
				<Table cellpadding="2" cellspacing="0" border="0" width="100%">
					<TR>
						<TD align="right" width="30%">
							Infantry
						</TD>
						<TD widt"20%">
							<% =formatnumber(rsspies("infantry"),0) %>
						</TD>
						<TD align="right" width="30%">
							Commandos
						</TD>
						<TD width="20%">
							<% =formatnumber(rsspies("comandoes"),0) %>
						</TD>
					</TR>
					<TR>
						<TD align="right">
							Jeeps
						</TD>
						<TD>
							<% =formatnumber(rsspies("jeeps"),0) %>
						</TD>
						<TD align="right">
							Matilda MkI
						</TD>
						<TD>
							<% =formatnumber(rsspies("matilda1"),0) %>
						</TD>
					</TR>
					<TR>
						<TD align="right">
							Matilda MkII
						</TD>
						<TD>
							<% =formatnumber(rsspies("matilda2"),0) %>
						</TD>
						<TD align="right">
							Sherman Tank
						</TD>
						<TD>
							<% =formatnumber(rsspies("sherman"),0) %>
						</TD>
					</TR>
					<TR>
						<TD align="right">
							Spitfire MkI
						</TD>
						<TD>
							<% =formatnumber(rsspies("spitfire1"),0) %>
						</TD>
						<TD align="right">
							Spitfire MkII
						</TD>
						<TD>
							<% =formatnumber(rsspies("Spitfire2"),0) %>
						</TD>
					</TR>
					<TR>
						<TD align="right">
							Spitfire MkV
						</TD>
						<TD>
							<% =formatnumber(rsspies("spitfire5"),0) %>
						</TD>
						<TD align="right">
							Hawker Hurricane
						</TD>
						<TD>
							<% =formatnumber(rsspies("hurricane"),0) %>
						</TD>
					</TR>
					<TR>
						<TD align="right">
							Halifax Bomber
						</TD>
						<TD>
							<% =formatnumber(rsspies("halifax"),0) %>
						</TD>
						<TD align="right">
							Wellington Bomber
						</TD>
						<TD>
							<% =formatnumber(rsspies("wellington"),0) %>
						</TD>
					</TR>
					<TR>
						<TD align="right">
							Lancaster Bomber
						</TD>
						<TD>
							<% =formatnumber(rsspies("lancaster"),0) %>
						</TD>
						<TD align="right">
							Landing Craft
						</TD>
						<TD>
							<% =formatnumber(rsspies("landingcraft"),0) %>
						</TD>
					</TR>
					<TR>
						<TD align="right">
							Landing Ships
						</TD>
						<TD>
							<% =formatnumber(rsspies("landingship"),0) %>
						</TD>
						<TD align="right">
							Frigates
						</TD>
						<TD>
							<% =formatnumber(rsspies("frigates"),0) %>
						</TD>
					</TR>
					<TR>
						<TD align="right">
							Battleships
						</TD>
						<TD>
							<% =formatnumber(rsspies("battleships"),0) %>
						</TD>
						<TD align="right">
							Aircraft Carriers
						</TD>
						<TD>
							<% =formatnumber(rsspies("carrier"),0) %>
						</TD>
					</TR>
					<%
					if cint(rsspies("type")) = 2 then
					'the stuff for the report with land defences
					%>
					<TR>
						<TD align="right">
							AA Guns
						</TD>
						<TD>
							<% =formatnumber(rsspies("aaguns"),0) %>
						</TD>
						<TD align="right">
							Pillboxes
						</TD>
						<TD>
							<% =formatnumber(rsspies("pilboxes"),0) %>
						</TD>
					</TR>
					<TR>
						<TD align="right">
							Turrets
						</TD>
						<TD>
							<% =formatnumber(rsspies("turrets"),0) %>
						</TD>
						<TD align="right">
							Barrage Balloons
						</TD>
						<TD>
							<% =formatnumber(rsspies("balloons"),0) %>
						</TD>
					</TR>
					<TR>
						<TD align="right">
							Water Mines
						</TD>
						<TD>
							<% =formatnumber(rsspies("mines"),0) %>
						</TD>
						<TD align="right">
							Spies
						</TD>
						<TD>
							<% =formatnumber(rsspies("spies"),0) %>
						</TD>
					</TR>
					<TR>
						<TD align="right">
							Military Police
						</TD>
						<TD>
							<% =formatnumber(rsspies("mp"),0) %>
						</TD>
					</TR>
					<%
					end if
					%>
					
				</table>
				<BR />
				<BR />
				<A HREF="delete_spy_report.asp?id=<% =lngreportid %>" onclick="return confirm('Are you sure you want to delete this report?');">Delete</A>
			</TD>
		</TR>
			
	</table>
</center>

<%
end if

rsspies.Close 

set rsspies = nothing
%>
<!---#include file="includes\footer.asp"--->
<!---#include file="includes\end_check.asp"--->
<%
end if
%>