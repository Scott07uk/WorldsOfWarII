<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
%>
<!---#include file="includes\top.asp"--->
<center>
<%
Dim RsReport
Dim strType

strType = Request.QueryString("type")

set rsreport = server.CreateObject("ADODB.Recordset")

select case strType
	case "abuse"
		strsql = "SELECT reportid, type, username, reportedlocation, detials FROM tblabusereports;"
	case "complaint"
		strsql = "SELECT reportid, type, username, reportedlocation, details FROM tblcheetingreports;"
	case "suggestion"
		strsql = "SELECT suggestionid, partfor, email, title, details FROM tblsuggestions;"
end case

RsReport.Open strsql, adocon

if RsReport.EOF then
%>
There are currently no reports to read.  
<%
else
do while not RsReport.EOF 
%>
<table cellpadding="2" border="0" cellspacing="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color@#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline" align="center">
			<% if (strType = "abuse") then %>
				Abuse report about location <% =rsreport("reportedlocation") %>
			<% elseif (strType = "complaint") then %>
				Complaint about location <% =rsreport("reportedlocation") %>
			<% else %>
				Suggestion: <% =rsreport("title") %>
			<% end if %>
		</TD>
	</TR>
	<TR>
		<TD class="norm">
			<table cellpadding="2" border="0" cellspacing="0">
				<TR>
					<TD class="norm" align="right">
						Type
					</TD>
					<TD class="norm">
						<% 
						if (strtype = "abuse") OR (strtype="complaint") then
								Response.Write(rsreport("type"))
							else
								Response.Write(rsreport("partfor"))
						end if
						%>
					</TD>
				</TR>>
					<TD class="norm">
						
					
			


<!---#include file="includes\footer.asp"--->

<%
end if
%>
<!---#include file="includes\end_check.asp"--->