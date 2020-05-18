<!---#include file="includes/check.asp"--->
<!---#include file="includes/top.asp"--->

<!---SPELL CHECKED--->
<center>
<table width="95%" cellpadding="0" border="0" cellspacing="0" style="border-top:none; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD align="center" class="norm">
			<A HREF="javascript:openHelp('16');">Help With This Page</A>
		</TD>
	</TR>
</table>
</center>
<%
Dim rsvote

set rsvote = server.CreateObject("ADODB.Recordset")

rsvote.LockType = 3
rsvote.CursorType = 2

strsql = "SELECT powerbase FROM tblaccounts WHERE continent = " & lngcontinentno & " ORDER BY numvotes DESC;"

rsvote.Open strsql, adocon

if not rsvote.EOF then
		call rsvote.Update("powerbase", cbool(true))
		rsvote.MoveNext 
end if

do while not rsvote.EOF 
	if cbool(rsvote("powerbase")) = true then
			call rsvote.Update("powerbase", cbool(false))
	end if
	rsvote.MoveNext 
loop

rsvote.Close
%>
<br />
<BR />
<center>
<%
select case Request.QueryString("msg")
	case 1
		Response.Write("Your vote has been updated.<BR /><BR />")
	case 2
		Response.Write("Your continent settings have been updated.<Br /><BR />")
end select
%>
</center>
<table align="center" cellpadding="2" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<tr>
		<td class="tdheaderline" width="70%" style="border-right: solid; border-width:1;">
			Candidate
		</td>
		<td class="tdheaderline" width="25%" style="border-right: solid; border-width:1;">
			Current Votes
		</td>
		<td class="tdheaderline" width="5%" style="border-right: solid; border-width:1;">&nbsp;
			
		</td>		
	</tr>
<%
'open the data base to see if the user is already voting for someone
strsql = "SELECT vote FROM tblaccounts WHERE username = '" & strusername & "';"

rsvote.Open strsql, adocon

if cstr(rsvote("vote")) <> "" then
		'the user does have an entry in their vote field make sure its valid
		strsql = "SELECT leadername, countryname FROM tblaccounts WHERE username = '" & cstr(rsvote("vote")) & "';"
		
		rsvote.Close 
		
		rsvote.Open strsql, adocon
		'their vote is for someone who has been deleted
		if not rsvote.EOF then
		'dispal who they are voting for
		%>
		<tr>
			<TD colspan="3" class="norm" align="center" style="border-bottom:solid; border-width:1; border-color:#FFFFFF;">
				You are currently voting for <% =rsvote("leadername") %> of <% =rsvote("countryname") %>
			</TD>
		</TR>
		<%
		end if
end if

rsvote.Close

StrSQL = "SELECT TblAccounts.username, continent, leadername, countryname, numvotes, powerbase FROM TblAccounts WHERE continent = " & lngContinentno & " ORDER BY numvotes DESC;"

rsvote.Open StrSQL, Adocon
do while not rsvote.eof
%>
	<tr>
		<td class="norm" style="border-right: solid; border-width:1;">
			<% if cbool(rsvote("powerbase")) = true then %>
			<font color="#FF0000">
			<% end if %>
			<%=rsvote("leadername")%> of <%=rsvote("countryname")%>
			<% if cbool(rsvote("powerbase")) = true then %>
			</font>
			<% end if %>
		</td>
		<td class="norm" style="border-right: solid; border-width:1;">
			<img SRC="images/votebar.gif" border="0" height="10" width="<% =clng(rsvote("numvotes"))*10 %>" /> &nbsp;
		</td>
		<td class="norm">
			<a href="vote.asp?username=<%=rsvote("username") %>">Vote</a>
		</TD>
	</TR>
<%
rsvote.MoveNext
loop

rsvote.Close 
%>
</table>
<%
strsql = "SELECT powerbase FROM tblaccounts WHERE username = '" & strusername & "';"

rsvote.Open strsql, adocon
Dim blnPowerBase
blnPowerBase = cbool(rsvote("powerbase"))


if blnpowerbase then

rsvote.Close 

strsql = "SELECT name, continentno, banner, message, ircchannel FROM tblcontinents WHERE continentno = " & lngcontinentno

rsvote.Open strsql, adocon

if rsvote.EOF then
		with rsvote
			.addnew
			.fields("name") = "[Worlds of War II]"
			.fields("continentno") = lngcontinentno
			.fields("message") = "No message has been set."
			.update
		end with
end if
%>
<BR />
<BR />
<center>
<table cellpadding="2" width="95%" border="0" cellspacing="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline" align="center">
			Power Base Options
		</TD>
	</TR>
	<TR>
		<TD class="norm" align="center">
			As the power base, you can edit your continent settings. 
			<BR />
			<BR />
			<form action="change_pb_settings.asp" method="post">
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
				</table>
				<BR />
				<input type="submit" value="Update Details" class="form">
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
<!---#include file="includes/footer.asp"--->
<!---#include file="includes/end_check.asp"--->