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
				Remove Suspension
			</TD>
		</TR>
		<TR>
			<TD class="norm" align="center">
				<%
				'declair varibles to be used
				Dim rsSuspend
				
				'generate the sql script
				strSQL = "SELECT username, continent, countryid FROM tblaccounts WHERE suspended = 1;"
				
				'create the objects to access the collection
				set rsSuspend = server.CreateObject("ADODB.Recordset")
				
				rsSuspend.Open strsql, adocon
				
				'check to see incase there are not suspended accounts
				if rsSuspend.EOF then
				%>
				There are no accounts suspended at the moment.
				<%
				else
				%>
				<table cellpadding="5" border="0" cellspacing="0" style="border-top:solid;border-bottom:solid;border-right:solid;border-left:solid;border-width:1;border-color:#FFFFFF;">
					<TR>
						<TD class="tdheaderline">
							Username
						</TD>
						<TD class="tdheaderline">
							Location
						</TD>
						<TD class="tdheaderline">
							Remove?
						</TD>
					</TR>
					<%
					do while not rsSuspend.EOF
					%>
					<TR>
						<TD class="norm">
							<% =rssuspend("username") %>
						</TD>
						<TD class="norm">
							<% =rssuspend("continent") %>
							:
							<% =rssuspend("countryid") %>
						</TD>
						<TD class="norm">
							<A HREF="admin_remove_suspention_engine.asp?username=<% =rssuspend("username") %>">Remove</A>
						</TD>
					</TR>
					<%
					rsSuspend.MoveNext
					loop
					%>
				</table>
				<%
				end if
				
				rsSuspend.Close
				
				set rssuspend = nothing
				%>
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