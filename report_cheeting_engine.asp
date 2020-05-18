<%
'SPELL CHECKED

'declair the pages varibles
Dim adocon			'holds the connection to the sql server
Dim strSQL			'holds the query ran on the server
Dim rsReport		'holds the report collection
Dim strDetails		'holds the details of the complaint
Dim strLocation		'holds the location of the complaint
Dim strUsername		'holds the username of the report
Dim strRedit		'holds the url of a page redirection
Dim intReportID		'holds the number for the report id

'get the form submissions
strDetails = Request.Form("description")
strLocation = Request.Form("continent") & ":" & Request.Form("country")
strUsername = Request.Form("username")

'tidy up the submission
strDetails = server.HTMLEncode(trim(strDetails))
strLocation = server.HTMLEncode(trim(strLocation))
strUsername = server.HTMLEncode(trim(strusername))

'generate a redirection varible
strRedir = "report_form.asp?type=cheet&username=" & strusername & "&con=" & Request.Form("continent") & "&country=" & Request.Form("country") & "&details=" & strdetails

'validate the user input
if strDetails = "" then
		'the user has left the explaintion field blank
		Response.Redirect(strRedir)
	else
		if strLocation = "" then
				Response.Redirect(strRedit)
			else
				'the sumbission is ok
				
				'create the objects needed to insert the record
				set adocon = server.CreateObject("ADODB.Connection")
				
				set rsReport = server.CreateObject("ADODB.Recordset")
				
				'generate the sql to query the database
				strsql = "SELECT * FROM tblcheetingReports ORDER BY ReportID DESC;"
				
				'open the connection
				
				adocon.ConnectionString = "Provider=SQLOLEDB;Server=neptune;User ID=WoWUsr;Password=password;Database=WoWII;"
				
				adocon.Open
				
				'open the recordset
				rsReport.Open strsql, adocon
				
				'check to see what the next report id will be
				
				if rsReport.EOF then
						intReportID = 1 
					else
						intReportID = cint(rsReport("ReportID") + 1)
				end if
				
				'save the report
				rsReport.AddNew 
				rsReport.Fields("reportid") = cint(intreportid)
				rsReport.Fields("type") = Request.Form("type")
				rsReport.Fields("username") = strusername
				rsReport.Fields("ReportedLocation") = strlocation
				rsReport.Fields("details") = strdetails
				rsReport.Fields("unread") = cbool(true)
				rsReport.Fields("replied") = cbool(false)
				
				'comence the changes
				rsReport.Update
				
				'tidy up our objects
				rsReport.Close
				adocon.Close
				
				set rsreport = nothing
				set adocon = nothing
				
				'redirect the user to the done page
				Response.Redirect("report_done.asp?reportid=c" & intreportid)
		end if
		
end if

%>