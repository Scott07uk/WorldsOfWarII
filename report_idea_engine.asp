<%
'SPELL CHECKED

Dim adocon
Dim strSQL
Dim rsIdea
Dim strType
Dim strEmail
Dim strTitle
Dim strdescription
Dim lngsuggestionid

'get the user input
strtype = Request.Form("type")
stremail = Request.Form("email")
strtitle = Request.Form("title")
strdescription = Request.Form("description")

'format the input

strtitle = trim(server.HTMLEncode(strtitle))
strdescription = trim(server.HTMLEncode(strdescription))

if (strtitle="") OR (strdescription ="") then
		Response.Redirect("report_form?type=suggestion")
	else
		'create the objects to do the work
		set adocon = server.CreateObject("ADODB.Connection")
		set rsidea = server.CreateObject("ADODB.recordset")
		
		'set the object attributes
		adocon.ConnectionString = "Provider=SQLOLEDB;Server=neptune;User ID=WoWUsr;Password=password;Database=WoWII;"
		
		rsIdea.LockType = 3 
		rsIdea.CursorLocation = 2
		
		'open the connection
		adocon.Open 
		
		'generate the sql
		strsql = "SELECT tblsuggestions.* FROM tblsuggestions ORDER by suggestionID DESC;"
		
		rsIdea.Open strsql, adocon
		
		if rsIdea.EOF then
				lngsuggestionid = 1
			else
				lngsuggestionid = clng(rsidea("suggestionid"))
				lngsuggestionid = lngsuggestionid + 1
		end if
		
		with rsidea
			.AddNew
			.Fields("suggestionid") = lngsuggestionid
			.Fields("partfor")= strtype
			.Fields("email") = stremail
			.Fields("title") = strtitle
			.Fields("details") = strdetails
			.Fields("unread") = cbool(true)
			
			.Update
		end with
		
		rsIdea.Close 
		adocon.Close 
		
		set rsidea = nothing
		set adocon = nothing
		
		Response.Redirect("report_done.asp?reportid=i" & lngsuggestionid)
end if
%>