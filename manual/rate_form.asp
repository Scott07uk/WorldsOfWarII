<%
if (isnumeric(Request.QueryString("id")) = false) OR (isnumeric(Request.QueryString("value")) = false) then
		Response.Write ("<SCRIPT language=""javascript"">")
		Response.Write("window.close();")
		Response.Write("</script>")
	else
		if (Request.QueryString("value") < 0) OR (Request.QueryString("value") > 5) then
				Response.Write ("<SCRIPT language=""javascript"">")
				Response.Write("window.close();")
				Response.Write("</script>")
			else

				Dim adocon
				Dim rsManual
				Dim strSQL

				'create the objects
				set adocon = server.CreateObject("ADODB.Connection")
				set rsManual = server.CreateObject("ADODB.Recordset")

				'set the attributes
				adocon.ConnectionString = "Provider=SQLOLEDB;Server=neptune;User ID=WoWUsr;Password=password;Database=WoWII;"
				
				strsql = "SELECT totalrating, numrating FROM tblmanual WHERE enteryid = " & Request.QueryString("id")
				
				adocon.Open 
				
				rsManual.LockType = 3
				rsManual.CursorType = 2
				
				rsManual.Open strsql, adocon
				
				if rsManual.EOF then
						Response.Write ("<SCRIPT language=""javascript"">")
						Response.Write("window.close();")
						Response.Write("</script>")
						rsManual.Close 
					else
						rsManual.Fields("totalrating") = clng(rsManual("totalrating")) + Request.QueryString("value")
						rsManual.Fields("numrating") = clng(rsmanual("numrating"))+1
						
						rsManual.Update 
						
						
						rsManual.Close
						%>
						<HTML>
	<HEAD>
		<TITLE>
			Worlds of War II Manual
		</TITLE>
		<link rel=stylesheet HREF="/includes/all.css">
	</HEAD>
	<BR />
	<Body bgcolor="black" link="white" alink="white" vlink="white" text="#ffffff" topmargin="0" bottommargin="0" leftmargin="0" rightmargin="0">
	<table cellpadding="2" cellspacing="0" border="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
		<TR>
			<TD class="tdheaderline" align="center">
				Manual Page Rateing
			</TD>
		</TR>
		<TR>
			<TD class="norm" align="center">
				Thank you for giving up feedback on the manual page, if you have any comments to make please leave them below.  
				<BR />
				<BR />
				<form action="rate_engine.asp" method="post">
					<input type="hidden" name="id" value="<% =Request.QueryString("ID") %>" />
					<textarea class="form" cols="60" rows="10" name="feedback"></TEXTAREA>
					<BR />
					<input type="submit" value="Leave Feedback" class="form" />
				</form>
				<BR />
				<BR />
				<A HREF="javascript:window.close();">Close Window</A>
			</TD>
		</tR>
	</table>
	</body>
</html>
						<%
				end if
				
				 
				adocon.Close 
				
				set rsManual = nothing
				set adocon = nothing
		end if
end if
%>