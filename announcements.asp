<!---#include file="includes/ptop.asp"--->

<!---SPELL CHECKED--->

<table cellpadding="5" border="0" cellspacing="5" style="border-left:solid; border-right:solid; border-top:solid; border-bottom:solid; border-width:1; border-color:#FFFFFF;" width="95%">
	<TR>
		<TD class="tdheader" align="center">
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
		Response.Write("<BR /><BR /><small><I>Posted by " & rsannounce("username") & " on " & rsannounce("posted"))
		end if
		rsannounce.Close
		
		set rsaccnounce = nothing
		adocon.Close
		set adocon = nothing
		%>
		</TD>
	</TR>
</table>
	

<!---#include file="includes/pfooter.asp"--->