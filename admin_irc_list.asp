<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
%>
<!---#include file="includes\top.asp"--->

<BR />
<BR />
<center>

<table cellpadding="2" width="95%" border="0" cellspacing="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline" align="center">
			Manage IRC Servers
		</TD>
	</TR>
	<TR>
		<TD class="norm" align="center">
			You can use this page to manage the IRC servers on the network.  
			<br />
			<b>DO NOT LINK ANY SERVER WITHOUT ADDING IT HERE</b>
			<br />
			<b>DO NOT UNLINK ANY SERVER WITHOUT REMOVING IT FROM HERE</b>
			<br />
			<b>DO NOT MAKE ANY CHANGES HERE IN THE EVENT OF A NET SPLIT, THE SERVERS ARE PROGRAMMED TO RECOVER AUTOMATICALLY</b>
			<br />
			<b>IF IN DOUBT DO <u>NOT</u> CHANGE ANY OF THESE SETTINGS</b>
			<BR />
			<BR />
			<table cellpadding="2" border="0" cellspacing="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;" width="95%">
				<TR>
					<TD class="tdheaderline">
						Numeric
					</TD>
					<TD class="tdheaderline">
						Server Name
					</TD>
					<TD class="tdheaderline">
						Managed By
					</TD>
					<TD class="tdheaderline">
						Connection Allowed
					</TD>
					<TD class="tdheaderline">
						Services
					</TD>
					<TD class="tdheaderline">
						Statistics
					</TD>
					<TD class="tdheaderline">
						Status
					</TD>
					<TD class="tdheaderline">
						Edit?
					</TD>
					<TD class="tdheaderline">
						De-Link?
					</TD>
				</TR>
				<%
				Dim rsIRC
				
				set rsirc = server.CreateObject("ADODB.Recordset")
				
				strsql = "SELECT ServerNumeric, AllowConnect, NetworkBots, StatsBot, OwnedBy, ContactEmail, Address, status FROM tblircservers;"
				
				rsIRC.Open strsql, adocon
				
				do while not rsIRC.EOF 
				%>
				<TR>
					<TD class="norm">
						<% =rsirc("servernumeric") %>
					</TD>
					<TD class="norm">
						<% =rsirc("address") %>
					</TD>
					<TD class="norm">
						<A HREF="mailto:<% =rsirc("contactemail") %>"><% =rsirc("ownedby") %></A>
					</TD>
					<TD class="norm">
						<% if cbool(rsirc("allowconnect")) = true then %>
								Yes
							<% else %>
								No
						<% end if %>
					</TD>
					<TD class="norm">
						<% if cbool(rsirc("networkbots")) = true then %>
								Yes
							<% else %>
								No
						<% end if %>
					</TD>
					<TD class="norm">
						<% if cbool(rsirc("statsbot")) = true then %>
								Yes
							<% else %>
								No
						<% end if %>
					</TD>
					<TD class="norm">
						<% 
						select case rsirc("status")
							case 1
								Response.Write("Online")
							case 2
								Response.Write("Offline")
							case 3
								Response.Write("Un-Configured")
						end select
						%>
					</TD>
					<TD class="norm">
						<A HREF="admin_edit_irc_server_form.asp?numeric=<% =rsirc("servernumeric") %>">Edit</A>
					</TD>
					<TD class="norm">
						<A HREF="admin_delete_irc_server.asp?numeric=<% =rsirc("servernumeric") %>" onclick="return confirm('Are you sure you want to de-link this server? Remember it will take up to 10 minutes for the changes to propagate the network.');">De-Link</A>
					</TD>
				</TR>
				<%
				rsIRC.MoveNext 
				loop
				
				rsIRC.Close
				set rsirc = nothing
				%>
			</table>
			<BR />
			<BR />
			<A Href="admin.asp">Main Admin Menu</A>
			
		</TR>
	</TR>
</table>
</center>

<!---#include file="includes\footer.asp"--->

<%
end if
%>
<!---#include file="includes\end_check.asp"--->