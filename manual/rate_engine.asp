<%
if (isnumeric(Request.Form("ID")) = false) OR (Request.Form("feedback") = "") then
		Response.Write ("<SCRIPT language=""javascript"">")
		Response.Write("window.close();")
		Response.Write("</script>") 
	else
		Dim rsManual
		Dim adocon
		Dim strSQL
		Dim lngID
		
		'create the objects
		set adocon = server.CreateObject("ADODB.Connection")
		set rsManual = server.CreateObject("ADODB.Recordset")

		'set the attributes
		adocon.ConnectionString = "Provider=SQLOLEDB;Server=neptune;User ID=WoWUsr;Password=password;Database=WoWII;"
				
		strSQL = "SELECT * FROM tblManualComments ORDER BY commentID DESC;"
				
		adocon.Open 
				
		rsManual.LockType = 3
		rsManual.CursorType = 2
				
		rsManual.Open strsql, adocon
				
		if rsManual.EOF then
				lngID = 1
			else
				lngID = clng(rsmanual("commentid")) + 1
		end if
				
		with rsmanual
			.AddNew
			.Fields("commentid") = lngID
			.Fields("manid") = Request.Form("id")
			.Fields("comment") = trim(server.HTMLEncode(Request.Form("feedback")))
					
			.Update
		end with
		
		rsManual.Close 
		
		adocon.Close 
		
		set rsmanual = nothing
		set adocon = nothing
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
	<center>
		<font face="arial" size="2">
		<BR />
		Thank you, your comments have been saved. 
		<BR /><BR />
		<A HREF="javascript:window.close();">Close Window</A>
	</font>
	</center>
	</body>
</html>
<%
end if
%>
				
				