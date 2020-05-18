<!---#include file="includes\check.asp"--->

<%
Dim rsusernamecheck		'varible to hold the username validation
Dim rsvotes				'recordset to do the work

'ok lets validate what the user has input
if (isnumeric(request.Form("option1")) = false) OR (isnumeric(request.Form("option2")) = false) OR (isnumeric(request.Form("option1")) = false) then
		'ther user has made their own form to send data to the page
		Response.Redirect("over.asp")
	else
		'ok lets make sure they havent made their own form to vote twice
		if (request.Form("option1") = request.Form("option2")) OR (request.Form("option2") = request.Form("option3")) OR (request.Form("option1") = request.Form("option3")) then
				'they are trying to vote for the same thing twice
				Response.Redirect("over.asp?bob")
			else
				'ok now lets check their username
				
				'(deano classes go in capitals)
				set rsusernamecheck = server.CreateObject("ADODB.recordset")

				strSQL = "SELECT * FROM tblUserVotes WHERE username = '" & strUsername & "';"

				'ok lets allow updating
				rsusernamecheck.LockType = 3
				rsusernamecheck.cursortype = 2

				'open her up
				rsusernamecheck.open strSQL, adocon

				'has the user already voted?
				if not rsusernamecheck.EOF then
						'yes they have they are being naughty
						Response.Redirect("over.asp")
					else
						'no they havent
						rsusernamecheck.AddNew
						rsusernamecheck.Fields("username") = strUsername
						rsusernamecheck.Update
						
						'we dont need this collection any more so lets close it NOW
						
						rsusernamecheck.Close 
						
						set rsvotes  = server.CreateObject("ADODB.recordset")
						
						'generate the sql
						strSQL = "SELECT ideano, votescore FROM tblideavotes WHERE"
						Dim counter
						'loop through all the records needed
						for counter = 1 to 3
							strsql = strsql & " ideano = " & request.Form("option" & counter)
							if counter < 3 then
									strsql = strsql & " OR"
							end if
						next

						'lets allow updating
						rsvotes.LockType = 3
						rsvotes.cursortype = 2
						
						'open her up
						rsvotes.open strSQL, adocon
					
						'ok lets loop through what we have got
						do while not rsVotes.EOF
							for counter = 1 to 3
								'lets see where the vote is in the list
								if clng(rsVotes("ideano")) = clng(request.Form("option" & counter)) then
										select case counter
											case 1
												rsvotes.fields("votescore") = clng(rsvotes("votescore")) + 3
											case 2
												rsvotes.fields("votescore") = clng(rsvotes("votescore")) + 2
											case 3
												rsvotes.fields("votescore") = clng(rsvotes("votescore")) + 1
										end select
										rsvotes.Update 
								end if
							next
							rsvotes.MoveNext 
						loop
						
						'ok it would appear all is done now lets tidy up
						
						rsvotes.Close 
						
						set rsvotes = nothing
						
						'send the user to the overview page
						response.Redirect("over.asp")
				end if
				
				'tidy up
				set rsusernamecheck = nothing
		end if
end if
%>

<!---#include file="includes\end_check.asp"--->