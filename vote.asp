<!---#include file="includes/check.asp"--->

<!---SPELL CHECKED--->

<%
if Request.QueryString("username") = "" then
		Response.Redirect("politics.asp")
	else
	
		dim rsVote

		set rsvote = server.CreateObject("ADODB.Recordset")
		
		rsVote.LockType = 3
		rsVote.CursorType = 2
		
		strsql = "SELECT vote FROM tblaccounts WHERE username = '" & strusername & "';"
		
		rsVote.Open strsql, adocon
		
		if rsvote("vote") <> "" then
				strsql = "SELECT numvotes FROM tblaccounts WHERE username = '" & rsvote("vote") & "';"
				
				rsVote.Close 
				
				rsVote.Open strsql, adocon
				
				if not rsVote.EOF then
						if rsvote("numvotes") < 2 then
								rsVote.Fields("numvotes") = 0
							else
								rsVote.Fields("numvotes") = cint(rsvote("numvotes")) - 1
						end if
						rsVote.Update 
				end if
		end if
		
		rsVote.Close 
		
		strsql = "SELECT numvotes FROM tblaccounts WHERE username = '" & Request.QueryString("username") & "';"
		
		rsVote.Open strsql, adocon
		if rsVote.EOF then
				Response.Redirect("politics.asp")
			else
				if rsvote("numvotes") < 0 then
						rsVote.Fields("numvotes") = 1
					else
						rsVote.Fields("numvotes") = cint(rsvote("numvotes")) + 1
				end if
				rsVote.Update 
				
		end if
		rsVote.Close 
				
		strSQL = "SELECT vote FROM tblaccounts WHERE username = '" & strusername & "';"
			
		rsVote.Open strsql, adocon
				
		rsVote.Fields("vote") = Request.QueryString("username")
				
		rsVote.Update 
		
		rsVote.Close
		
		set rsvote = nothing
		Response.Redirect("politics.asp?msg=1")
end if
					
%>
<!---#include file="includes/end_check.asp"--->