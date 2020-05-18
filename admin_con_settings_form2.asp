<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
%>
<!---#include file="includes\top.asp"--->


<%
Dim rsvote

set rsvote = server.CreateObject("ADODB.Recordset")

strsql = "SELECT name, continentno, banner, message, ircchannel, locked FROM tblcontinents WHERE continentno = " & Request.Form("continent")

rsvote.Open strsql, adocon

if rsvote.EOF then

%>
<BR /><BR />
<center>
	<font face="arial" size="2">
		The continent number you selected is invalid.
		<BR /><BR />
		Click <A HREF="admin.asp">here</A> to go back to the admin menu.
	</font>
</center>
<%
else
%>
<BR />
<BR />
<center>
<table cellpadding="2" width="95%" border="0" cellspacing="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline" align="center">
			Change Continent Settings
		</TD>
	</TR>
	<TR>
		<TD class="norm" align="center">
			Here you can change the options for your continent. 
			<BR />
			<BR />
			<form action="admin_con_settings_engine.asp" method="post">
				<input type="hidden" name="continent" value="<% =Request.Form("continent") %>" />
				<table cellpadding="2" cellspacing="0" border="0">
					<TR>
						<TD class="norm" align="right">
							Continent Name
						</TD>
						<TD class="norm">
							<input type="text" size="40" name="name" maxlength="50" class="form" value="<% =rsvote("name") %>" />
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right">
							Continent Banner
						</TD>
						<TD class="norm">
							<input type="text" size="40" name="banner" maxlength="100" class="form" value="<% =rsvote("banner") %>" />
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right">
							Continent IRC Channel
						</TD>
						<TD class="norm">
							<input type="text" size="25" name="ircchannel" maxlength="50" class="form" value="<% =rsvote("ircchannel") %>" />
						</TD>
					</TR>
					<TR>
						<TD class="norm" valign="top" align="right">
							Message from Commander
						</TD>
						<TD class="norm">
							<%
							Dim strmessage
							
							strmessage = rsvote("message")
							
							strmessage = replace(strmessage,"<BR>", vbcrlf)
							strmessage = replace(strmessage,"<B>", [B])
							strmessage = replace(strmessage,"</B>", "[/B]")
							strmessage = replace(strmessage,"<I>", [I])
							strmessage = replace(strmessage,"</I>", "[/I]")
							strmessage = replace(strmessage,"<U>", [U])
							strmessage = replace(strmessage,"</U>", "[/U]")
							%>
							<textarea name="message" cols="50" rows="6" class="form"><% =strmessage %></textarea>
						</TD>
					</TR>
					<TR>
						<TD class="norm" valign="top" align="right">
							Continent Locked:
						</TD>
						<TD class="norm">
						<%
						if cbool(rsvote("locked")) = true then 
						%>
							<input type="checkbox" name="locked" value="true" checked />
						<%
						else
						%>
							<input type="checkbox" name="locked" value="true" />
						<% 
						end if
						%>
					</TD>
				</TR>
				</table>
				<BR />
				<input type="submit" value="Update Details" class="form" />
			</form>
		</TD>
	</TR>
</table>
</center>
<%
end if
rsvote.Close 
set rsvote = nothing
%>

<!---#include file="includes\footer.asp"--->
<%
end if
%>
<!---#include file="includes\end_check.asp"--->