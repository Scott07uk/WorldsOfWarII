<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
%>
<!---#include file="includes\top.asp"--->
<%
dim rsconfig

strsql = "SELECT gameaddress, portaladdress, forumsaddress, manualaddress, ircsite FROM tblconfig;"

set rsconfig = server.CreateObject("ADODB.Recordset")

rsconfig.Open strsql, adocon
%>
<BR />
<BR />
<center>
<table cellpadding="2" width="95%" border="0" cellspacing="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline" align="center">
			Change Game URLs
		</TD>
	</TR>
	<TR>
		<TD class="norm" align="center">
			You can use this page to change the URLs that the game uses to point to the various sections of the game.  
			<BR />
			<BR />
			<form action="admin_update_url_engine.asp" method="post">
				<table cellpadding="2" border="0" cellpadding="0">
					<TR>
						<TD class="tdheader" align="right">
							Main Game URL
						</TD>
						<TD>
							<input type="text" size="35" maxlength="200" class="form" value="<% =rsconfig("gameaddress") %>" name="game" />
						</TD>
					</TR>
					<TR>
						<TD class="tdheader" align="right">
							Portal URL
						</TD>
						<TD class="norm">
							<input type="text" size="35" maxlength="200" class="form" value="<% =rsconfig("portaladdress") %>" name="portal" />
						</TD>
					</TR>
					<TR>
						<TD class="tdheader" align="right">
							Forums URL
						</TD>
						<TD class="norm">
							<input type="text" size="35" maxlength="200" class="form" value="<% =rsconfig("forumsaddress") %>" name="forums" />
						</TD>
					</TR>
					<TR>
						<TD class="tdheader" align="right">
							Manual URL
						</TD>
						<TD class="norm">
							<input type="text" size="35" maxlength="200" class="form" value="<% =rsconfig("manualaddress") %>" name="manual" />
						</TD>
					</TR>
					<TR>
						<TD class="tdheader" align="right">
							IRC Site
						</TD>
						<TD class="norm">
							<input type="text" size="35" maxlength="200" class="form" value="<% =rsconfig("ircsite") %>" name="irc" />
						</TD>
					</TR>
				</table>
				<BR />
				<input type="submit" value="Update" class="form">
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
%>
<!---#include file="includes\end_check.asp"--->