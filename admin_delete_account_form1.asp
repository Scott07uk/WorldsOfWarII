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
	<table cellpadding="2" border="0" cellspacing="0" width="95%" style="border-top:solid; border-right:solid; border-left:solid; border-bottom:solid; border-width:1; border-color:#FFFFFF;">
		<TR>
			<TD class="tdheaderline" align="center">
				Delete Account
			</TD>
		</TR>
		<TR>
			<TD class="norm" align="center">
				Use this page to view account details and delete them.  
				<BR />
				<BR />
				<Table cellpadding="2" cellspacing="0" border="0">
					<TR>
						<TD class="tdheaderline">
							Location
						</TD>
						<TD class="tdheaderline">
							Username
						</TD>
						<TD class="tdheaderline">
							Leader of Country
						</TD>
						<TD class="tdheaderline">
							Suspended
						</TD>
						<TD class="tdheaderline">
							Last Login
						</TD>
						<TD class="tdheaderline">
							Delete?
						</TD>
					</TR>
					<%
					dim RsDel
					
					set rsDel = server.CreateObject("ADODB.Recordset")
					
					strsql = "SELECT continent, countryid, username, leadername, countryname, suspended, lastlogin FROM tblaccounts;"
					
					RsDel.Open strsql, adocon
					do while not RsDel.EOF 
					%>
					<TR>
						<TD class="norm">
							<% =rsdel("continent") %>:<% =rsdel("countryid") %>
						</TD>
						<TD class="norm">
							<% =rsdel("username") %>
						</TD>
						<TD class="norm">
							<% =rsdel("leadername") %> of <% =rsdel("countryname") %>
						<TD class="norm">
							<% if cbool(rsdel("suspended")) = true then %>
								Yes
							<% else %>
								No
							<% end if %>
						</TD>
						<TD class="norm">
							<% =rsdel("lastlogin") %>
						</TD>
						<TD class="norm">
							<A onclick="return confirm('Are you sure you want to delete this account?')" HREF="admin_del_account_engine.asp?username=<% =rsdel("username") %>">Delete</A>
						</TD>
					</TR>
					<%
					RsDel.MoveNext 
					loop
					
					RsDel.Close 
					set rsdel = nothing
					%>
					</tablE>
				<BR />
				<BR />
				<A Href="admin.asp">Main Admin Menu</A>
			</TD>
		</TR>
	</table>
</center>

<!---#include file="includes\footer.asp"--->

<%
end if
%>
<!---#include file="includes\end_check.asp"--->