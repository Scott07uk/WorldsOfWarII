<!---#include file="includes/ptop.asp"--->

<!---SPELL CHECKED--->

<img src="image.asp?name=logo.gif" border="0" alt="Worlds of War II" />
<BR />
<BR />
<table cellpadding="2" cellspacing="0" border="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
	<TR>
		<TD class="tdheaderline" align="center">
			<font size="3" color="#FF0000">
				Welcome to Worlds of War II
			</font>
		</TD>
	</TR>
	<TR>
		<TD class="norm">
			Worlds of War II is the sequel to the original popular browser based strategy game. It is based around the British units of the Second World War. After signing up, you will find yourself placed in a random continent where you can work with or against your continent neighbours to try and dominate the game with a strong military and defence network. Take control of enemy lands and gold supplies using a wide array of land, sea and air units.
			<BR />
			<BR />
			If you're interested, why not <A HREF="register1.asp">sign up</A> right now, or join #worldsofwar on <A HREF="irc://irc.worldsofwar.net:6667">irc.worldsofwar.net</A> or <A HREF="irc://irc2.worldsofwar.net">irc2.worldosfwar.net</A> and meet the players and creators.
		</TD>
	</TR>
</table>
<BR />	
<BR />
<table cellpadding="2" border="0" cellspacing="0" style="border-left:solid; border-right:solid; border-top:solid; border-bottom:solid; border-width:1; border-color:#FFFFFF;" width="95%">
	<TR>
		<TD class="tdheaderline" align="center">
			Latest Announcement
		</TD>
	</TR>
	<TR>
		<TD class="norm">
		<%
		Dim rsannounce
		
		Set adoCon = Server.CreateObject("ADODB.Connection")
		
		adocon.ConnectionString = "Provider=SQLOLEDB;Server=neptune;User ID=WoWUsr;Password=password;Database=WoWII;"
		
		adocon.Open 
		
		strsql = "SELECT Message, username, posted FROM tblannouncements ORDER BY AnnouncementID DESC;"
		
		set rsannounce = server.CreateObject("ADODB.Recordset")
		
		rsannounce.Open strsql, adocon
		
		if rsannounce.EOF then
		%>
			There are currently no announcements to read.
		<%
		else
		Response.Write(rsannounce("message"))
		Response.Write("<div align=""right""><BR /><small><I>Posted by " & rsannounce("username") & " on " & left(rsannounce("posted"),16) & "</div>")
		end if
		rsannounce.Close
		
		set rsaccnounce = nothing
		adocon.Close
		set adocon = nothing
		%>
		</TD>
	</TR>
</table>
<BR />
<BR />
<!---#include file="includes/pfooter.asp"--->