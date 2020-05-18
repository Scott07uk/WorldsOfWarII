<!---#include file="includes/check.asp"--->

<!---SPELL CHECKED--->

<%
Dim rsuser

set rsuser = server.createobject("ADODB.recordset")

strsql = "SELECT  ircnick from tblaccounts where username = '" & strusername & "';"

rsuser.open strsql, adocon

%>
<center>
<applet code=IRCApplet.class archive="irc.jar,pixx.jar" width=640 height=400>
	<param name="CABINETS" value="irc.cab,securedirc.cab,pixx.cab">
	<param name="nick" value="<% =rsuser("ircnick") %>">
	<param name="alternatenick" value="<% =rsuser("ircnick") %>01">
<%
rsuser.close

strsql = "SELECT address, estusers FROM tblircservers WHERE allowconnect=1 AND status = 1 ORDER BY estusers ASC;"

rsuser.locktype = 3
rsuser.cursortype = 2 

rsuser.open strsql, adocon

Dim strirc1
Dim strirc2
if rsuser.eof then

strirc1 = "irc.worldsofwar.co.uk"
stritc2 = "irc2.worldsofwar.co.uk"

else

call rsuser.update("estusers", clng(rsuser("estusers") + 1))

strirc1 = rsuser("address")
stritc2 = stritc1

rsuser.movenext
if not rsuser.eof then
	strirc2 = rsuser("address")
end if
end if

rsuser.close

strsql = "SELECT IRCChannel FROM TblContinents WHERE ContinentNo = "  & clng(lngcontinentno)

rsuser.open strsql, adocon

strcontinentchan = rsuser("IRCChannel")

rsuser.close

set rsuser = nothing

%>
	<param name="fullname" value="PJIRC User">
	<param name="host" value="<% =strirc1 %>">
	<param name="gui" value="pixx">
	<param name="port" value="6667">
	<param name="command1" value="/join #worldsofwar">
<%
if strcontinentchan <> "" then
%>
	<param name="command2" value="/join #<% =strcontinentchan %>">
<%
end if
%>
	<param name="aternateserver1" value="<% =strirc2 %>">
	<param name="userid" value="<% =strusername %>">
	<param name="allowdccchat" value="false">
	<param name="allowdccfile" value="false">
		
</applet>
</center>
<!---#include file="includes\end_check.asp"--->