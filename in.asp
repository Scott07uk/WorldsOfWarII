<%
if trim(Request.QueryString("in")) <> "" then
		Dim adocon
		Dim rsAdd
			
		set adocon = server.CreateObject("ADODB.Connection")
			
		adocon.ConnectionString = "Provider=SQLOLEDB;Server=neptune;User ID=WoWUsr;Password=password;Database=WoWII;"
			
		adocon.Open 

		set rsAdd = server.CreateObject("ADODB.Recordset")

		rsAdd.LockType = 3

		rsAdd.CursorType = 2

		strsql = "SELECT * FROM tblInLinks WHERE site = '" & Request.QueryString("in") & "';"
		
		rsAdd.Open strsql, adocon
		
		if rsAdd.EOF then
				rsAdd.AddNew 
				rsAdd.Fields("site") = Request.QueryString("in")
				rsAdd.Fields("hits") = 1
			else
				rsAdd.Fields("hits") = clng(rsadd("hits")) + 1
		end if
		
		rsAdd.Update 
		
		rsAdd.Close
		
		adocon.Close
		
		set rsadd = nothing
		
		set adocon = nothing
end if

Response.Redirect("default.asp")
%>
		