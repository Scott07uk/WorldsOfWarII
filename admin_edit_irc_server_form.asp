<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else

Dim rsIRC
		
strsql = "SELECT ServerNumeric, AllowConnect, NetworkBots, StatsBot, OwnedBy, ContactEmail, Address, status FROM tblircservers WHERE servernumeric = " & Request.QueryString("numeric")
		
set rsIRC = server.CreateObject("ADODB.Recordset")
		
rsIRC.Open strsql, adocon
if rsIRC.EOF then
		Response.Redirect("admin_irc_list.asp")
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
			Edit the required settings below.  
			<br />
			<b>REMEMBER ALL CHANGES CAN TAKE UP TO 10 MINUTES TO PROPAGATE THE NETWORK</b>
			<br />
			<form action="admin_edit_irc_engine.asp" method="post">
				<table cellpadding="2" border="0" cellspacing="0">
					<TR>
						<TD class="norm" align="right">
							Server Numeric
						</TD>
						<TD class="norm">
							<input type="text" size="3" maxlength="3" READONLY name="numeric" value="<% =rsirc("servernumeric") %>" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right">
							Network Name
						</TD>
						<TD class="norm">
							<input type="text" size="40" maxlength="200" name="name" value="<% =rsirc("address") %>" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right">
							Server Admin
						</TD>
						<TD class="norm">
							<input type="text" size="15" maxlength="50" name="adminname" value="<% =rsirc("ownedby") %>" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right">
							Contact Email
						</TD>
						<TD class="norm">
							<input type="text" size="25" maxlength="150" name="adminemail" value="<% =rsirc("contactemail") %>" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right">
							Allow Connections?
						</TD>
						<TD class="norm">
							<SELECT name="connectallow" size="1" class="form">
								<% if cbool(rsirc("allowconnect")) = true then %>
								<option value="True" SELECTED>Yes</option>
								<option value="False">No</option>
								<% else %>
								<option value="True">Yes</option>
								<option value="False" SELECTED>No</option>
								<% end if %>
							</SELECT>
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right">
							Provides Services?
						</TD>
						<TD class="norm">
							<SELECT name="services" size="1" class="form">
								<% if cbool(rsirc("networkbots")) = true then %>
								<option value="True" SELECTED>Yes</option>
								<option value="False">No</option>
								<% else %>
								<option value="True">Yes</option>
								<option value="False" SELECTED>No</option>
								<% end if %>
							</SELECT>
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right">
							Provides Statistics?
						</TD>
						<TD class="norm">
							<SELECT name="stats" size="1" class="form">
								<% if cbool(rsirc("statsbot")) = true then %>
								<option value="True" SELECTED>Yes</option>
								<option value="False">No</option>
								<% else %>
								<option value="True">Yes</option>
								<option value="False" SELECTED>No</option>
								<% end if %>
							</SELECT>
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right">
							Status?
						</TD>
						<TD class="norm">
							<SELECT name="status" size="1" class="form">
								<% if rsirc("status") = 1 then %>
								<option value="1" SELECTED>Online</option>
								<% else %>
								<option value="1">Online</option>
								<% end if
								if rsirc("status") = 2 then %>
								<option value="2" SELECTED>Offline</option>
								<% else %>
								<option value="2">Offline</option>
								<% end if
								if rsirc("status") = 3 then %>
								<option value="3" SELECTED>De-Configured</option>
								<% else %>
								<option value="3">De-Configured</option>
								<% end if %>
							</SELECT>
						</TD>
					</TR>
				</table>
				<BR />
				<input type="submit" value="Update" class="form" />
			</form>
			
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
rsIRC.Close 

set rsirc = nothing
end if
%>
<!---#include file="includes\end_check.asp"--->