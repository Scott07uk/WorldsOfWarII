<HTML>
	<HEAD>
		<TITLE>
			Worlds of War II Manual
		</TITLE>
		<link rel=stylesheet HREF="/includes/all.css">
	</HEAD>
	<Body bgcolor="black" link="white" alink="white" vlink="white" text="#ffffff" topmargin="0" bottommargin="0" leftmargin="0" rightmargin="0">

	<table cellpadding="0" border="0" cellspacing="0" width="100%">
		<TR>
			<TD width="190" valign="top">
				<center>
					<IMG SRC="/images/logo.gif" alt="Worlds of War II" border="0" />
				</center>
				<BR />
				<table width="100%" border="0" cellpadding="2" cellspacing="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-width:1;border-color:#FFFFFF;">
					<TR>
						<TD class="tdheaderline" align="center">
							<A HREF="default.asp">Manual</A>
						</TD>
					</TR>
					<TR>
						<TD class="norm">
							<%
							Dim adocon
							Dim rsManual
							Dim strSQL
							
							'create the objects
							set adocon = server.CreateObject("ADODB.Connection")
							set rsManual = server.CreateObject("ADODB.Recordset")
							
							'set the attributes
							adocon.ConnectionString = "Provider=SQLOLEDB;Server=neptune;User ID=WoWUsr;Password=password;Database=WoWII;"
							
							'set the sql
							strSQL = "SELECT EnteryID, Entryname FROM TblManual WHERE OnSidebar = 1 ORDER BY Ordering ASC;"
							
							'open the recordset
							adocon.Open 
							
							rsManual.Open strsql, adocon
							
							do while not rsManual.EOF 
								Response.Write("<A HREF=""manual.asp?view=" & rsmanual("enteryid") & """>" & rsManual("entryname") & "</A>" & vbcrlf & "<BR />")
								rsManual.MoveNext 
							loop
							
							rsManual.Close
							
							%>
							<BR />
							<A HREF="contact.asp">Contact Us</A>
						</TD>
					</TR>
				</table>
			</TD>
			<TD valign="top" align="center">