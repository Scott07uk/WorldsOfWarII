<!---#include file="includes\check.asp"--->
<!---#include file="includes\top.asp"--->

<!---SPELL CHECKED--->
<center>
<table width="95%" cellpadding="0" border="0" cellspacing="0" style="border-top:none; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD align="center" class="norm">
			<A HREF="javascript:openHelp('3');">Help With This Page</A>
		</TD>
	</TR>
</table>
</center>
<center>
	<BR />
	
	<table cellpadding="2" border="0" width="95%" cellspacing="0" style="border-top:solid; border-left:solid; border-right:solid; border-bottom:solid; border-color:#FFFFFF; border-width:1;">
		<TR>
			<TD class="tdheaderline" align="center">
				Incoming Forces
			</TD>
		</TR>
		<TR>
			<TD class="norm" width="100%">
				<%
				Dim rsMovements		'collection of movement forces
				
				'create the database query
				strsql = "SELECT ETA, infantry, comandoes, jeeps, matilda1, matilda2, sherman, spitfire1, spitfire2, spitfire5, hurricane, lancaster, wellington, halifax, frigates, battleships, carrier, status, fromcountry, fromcontinent, tocountry FROM tblmovements WHERE tocontinent = " & lngcontinentno & " ORDER BY tocountry ASC;"
				
				
				set rsMovements = server.CreateObject("ADODB.Recordset")
				
				rsMovements.Open strsql, adocon
				
				if not rsMovements.EOF then
				%>
				<table cellpadding="0" cellspacing="2" border="0" width="100%">
					<TR>
						<TD class="norm" width="20%">
							Target
						</TD>
						<TD class="norm" width="20%"> 
							Sender
						</TD>
						<TD class="norm" width="20%">
							Units
						</TD>
						<TD class="norm" width="20%">
							ETA
						</TD>
						<TD class="norm" width="20%">
							Mission
						</TD>
					</TR>
					<%
					Dim strTextColor
					Dim strMissionName
					do while not rsMovements.EOF 
					if rsmovements("status") = 1 then
							strTextColor = "#FF0000"
							strMissionName = "Attack"
						else
							if rsmovements("status") = 2 then
									strTextColor = "#00FF00"
									strMissionName = "Defend"
								else
									strTextColor = "#AAAAAA"
									strMissionName = "Return"
							end if
					end if
					
					%>
					<TR>
						<TD class="norm">
							<font color="<% =strtextcolor %>">
								<% =lngcontinentno %>:<% =rsmovements("tocountry") %>
							</font>
						</TD>
						<TD class="norm">
							<font color="<% =strtextcolor %>">
								<% =rsmovements("fromcontinent") %>:<% =rsmovements("fromcountry") %>
							</font>
						</TD>
						<TD class="norm">
							<font color="<% =strtextcolor %>">
								<%
								Dim dblunitstotal		'varible to hold the total incomming units
								dblunitstotal = cdbl(rsmovements("infantry"))+cdbl(rsmovements("comandoes"))+cdbl(rsmovements("jeeps"))+cdbl(rsmovements("matilda1"))+cdbl(rsmovements("matilda2"))+cdbl(rsmovements("sherman"))+cdbl(rsmovements("spitfire1"))+cdbl(rsmovements("spitfire2"))+cdbl(rsmovements("spitfire5"))+cdbl(rsmovements("hurricane"))+cdbl(rsmovements("lancaster"))+cdbl(rsmovements("wellington"))+cdbl(rsmovements("halifax"))+cdbl(rsmovements("frigates"))+cdbl(rsmovements("battleships"))+cdbl(rsmovements("carrier"))
								%>
								<% =dblunitstotal %>
							</font>
						</TD>
						<TD class="norm">
							<font color="<% =strtextcolor %>">
								<% =rsmovements("eta") %>
							</font>
						</TD>
						<TD class="norm">
							<font color="<% =strtextcolor %>">
								<% =strmissionname %>
							</font>
						</TD>
					</TR>
					<%
					rsMovements.MoveNext
					loop
					end if
					rsMovements.Close
					%>
				</table>
				
			</TD>
		</TR>
	</table>
	
	<BR />
	<BR />
	
	<table cellpadding="2" border="0" width="95%" cellspacing="0" style="border-top:solid; border-left:solid; border-right:solid; border-bottom:solid; border-color:#FFFFFF; border-width:1;">
		<TR>
			<TD class="tdheaderline" align="center">
				Outgoing Forces
			</TD>
		</TR>
		<TR>
			<TD class="norm" width="100%">
				<%
				'create the database query
				strsql = "SELECT ETA, infantry, comandoes, jeeps, matilda1, matilda2, sherman, spitfire1, spitfire2, spitfire5, hurricane, lancaster, wellington, halifax, frigates, battleships, carrier, status, fromcountry, tocontinent, tocountry FROM tblmovements WHERE fromcontinent = " & lngcontinentno & " ORDER BY fromcountry ASC;"
				
				rsMovements.Open strsql, adocon
				
				if not rsMovements.EOF then
				%>
				<table cellpadding="0" cellspacing="2" border="0" width="100%">
					<TR>
						<TD class="norm" width="20%">
							Sender
						</TD>
						<TD class="norm" width="20%"> 
							Target
						</TD>
						<TD class="norm" width="20%">
							Units
						</TD>
						<TD class="norm" width="20%">
							ETA
						</TD>
						<TD class="norm" width="20%">
							Mission
						</TD>
					</TR>
					<%
					do while not rsMovements.EOF 
					if rsmovements("status") = 1 then
							strTextColor = "#FF0000"
							strMissionName = "Attack"
						else
							if rsmovements("status") = 2 then
									strTextColor = "#00FF00"
									strMissionName = "Defend"
								else
									strTextColor = "#AAAAAA"
									strMissionName = "Return"
							end if
					end if
					
					%>
					<TR>
						<TD class="norm">
							<font color="<% =strtextcolor %>">
								<% =lngcontinentno %>:<% =rsmovements("fromcountry") %>
							</font>
						</TD>
						<TD class="norm">
							<font color="<% =strtextcolor %>">
								<% =rsmovements("tocontinent") %>:<% =rsmovements("tocountry") %>
							</font>
						</TD>
						<TD class="norm">
							<font color="<% =strtextcolor %>">
								<%
								dblunitstotal = cdbl(rsmovements("infantry"))+cdbl(rsmovements("comandoes"))+cdbl(rsmovements("jeeps"))+cdbl(rsmovements("matilda1"))+cdbl(rsmovements("matilda2"))+cdbl(rsmovements("sherman"))+cdbl(rsmovements("spitfire1"))+cdbl(rsmovements("spitfire2"))+cdbl(rsmovements("spitfire5"))+cdbl(rsmovements("hurricane"))+cdbl(rsmovements("lancaster"))+cdbl(rsmovements("wellington"))+cdbl(rsmovements("halifax"))+cdbl(rsmovements("frigates"))+cdbl(rsmovements("battleships"))+cdbl(rsmovements("carrier"))
								%>
								<% =dblunitstotal %>
							</font>
						</TD>
						<TD class="norm">
							<font color="<% =strtextcolor %>">
								<% =rsmovements("eta") %>
							</font>
						</TD>
						<TD class="norm">
							<font color="<% =strtextcolor %>">
								<% =strmissionname %>
							</font>
						</TD>
					</TR>
					<%
					rsMovements.MoveNext
					loop
					end if
					rsMovements.Close
					%>
				</table>
				
			</TD>
		</TR>
	</table>
				

</center>

<!---#include file="includes\footer.asp"--->
<!---#include file="includes\end_check.asp"--->