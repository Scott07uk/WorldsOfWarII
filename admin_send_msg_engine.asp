<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else

'codes goes below here

'Declair varibles to be used
Dim rsMessage
Dim rsCheck
Dim strBody
Dim strSubject
Dim lngContinent
Dim lngCountry

'get the user input

strbody = Request.Form("body")
strsubject = Request.Form("subject")
lngContinent = Request.Form("continent")
lngcountry = Request.Form("country")

if (strbody = "") OR (strsubject = "") OR (isnumeric(lngcontinent) = false) then
		Response.Redirect("admin_send_notice_form.asp?error=1")
	else
		if lngCountry = "" then
				'the message is going to all countrys
				set rscheck = server.CreateObject("ADODB.Recordset")
				set RsMessage = server.CreateObject("ADODB.Recordset")
				
				rsMessage.LockType = 3
				rsMessage.CursorType = 2
				
				strsql = "SELECT countryid FROM tblaccounts WHERE continent = " & lngcontinent
				
				rsCheck.Open strsql, adocon
				
				if rsCheck.EOF then
						Response.Redirect("admin_send_notice_form.asp?error=2")
					else
						strsql = "SELECT tblmessages.* FROM tblmessages ORDER BY messageid DESC;"
						
						rsMessage.Open strsql, adocon
						
						Dim lngmessageid
						
						if rsmessage.eof then
								lngmessageid = 0
							else
								lngmessageid = clng(rsmessage("messageid"))
						end if
						
						do while not rsCheck.EOF
							lngmessageid = clng(lngmessageid + 1)
							with rsmessage
								.AddNew
								.Fields("tocontinent") = lngcontinent
								.Fields("tocountry") = rscheck("countryid")
								.Fields("fromcountry") = 0
								.Fields("fromcontinent") = 0
								.Fields("subject") = strsubject
								.Fields("body") = trim(strbody)
								.Fields("datesent") = now()
								.Fields("new") = 1
								.Fields("messageid") = lngmessageid
								
								.Update
							end with
							
							rsCheck.MoveNext 
						loop
						
						rsMessage.Close 
						rsCheck.Close 
						
						set rsmessage = nothing
						set rschec = nothing
					else
						if isnumeric(lngcountry) = false then
								Response.Redirect("admin_section_send_notice_form.asp?error=3")
							else
								set rsMessage = server.CreateObject("ADODB.Recordset")
								
								rsMessage.locktype = 3
								
								rsMessage.cursortype = 2
								
								strsql = "SELECT * FROM tblmessages ORDER BY messageID DESC;"
								
								rsmessage.open strsql, adocon
								
								if rsmessage.eof then
										lngmessageid = 1
									else
										lngmessageid = clng(rsmessage("messageid")) + 1
								end if
								
								with rsmessage
									.addnew
									.fields("tocountry") = lngcountry
									.fields("tocontinent") = lngcontinent
									.fields("fromcountry") = 0
									.fields("fromcontinent") = 0
									.fields("subject") = strsubject
									.fields("body") = strbody
									.fields("datesent") = now()
									.fields("new") = 1
									.fields("messageid") = lngmessageid
									
									.update
								end with
								
								rsmessage.close
								
								set rsmessage = nothing
						end if
				end if
		end if
end if
								

'code goes above here


end if
%>
<!---#include file="includes\end_check.asp"--->


