<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else

'code goes below here

'Declair varble for use in this page
Dim rsSuspend			'Recordset for suspending the user
Dim strSuspendName		'the username of the account to be suspended
Dim intCountry			'the country number of the account to be suspended
Dim intContinent		'the continent number of the account to be suspended
Dim strReason			'the reason for the account being suspended
Dim blnUseLoc			'the varible to decide if we are using username or location
Dim intError			'the error code
Dim objEmail			'the email object string to send the email

'get the input from the user
strSuspendName = Request.Form("username")
intContinent = Request.Form("continent")
intCountry = Request.Form("country")
strReason = Request.Form("reason")
intError = 0

'validate the information sumitted
if strSuspendName = "" then
		if (intCountry = "") or (intContinent = "") then
				intError=1
			else
				'we are using the location
				blnUseLoc = True
				'validate to make sure the numbers are correct
				if (isnumeric(intcountry)=False) or (isnumeric(intcontinent)=False) then
						interror = 2
					else
						'the submission is fine
				end if
		end if
	else
		if (intCountry <> "") or (intContinent <> "") then
				intError = 3
			else
				blnUseLoc=False
		end if
end if

if intError <> 0 then
		'there has been an error return them to the page
		Response.Redirect("admin_suspend_acc_form1.asp?error=" & intError)
	else
		'format the reason and validate it
		strReason = trim(strReason)
		if strReason = "" then
				Response.Redirect("admin_suspend_acc_form1.asp?error=4")
			else
				strReason = server.HTMLEncode(strReason)
				
				'generate the sql script to suspend the account
				strsql = "SELECT suspended, email, username FROM tblaccounts WHERE "
				
				'check to see how we are suspending the accounts
				if blnUseLoc = True then
						strsql = strsql & "continent = " & intcontinent & " AND countryid = " & intcountry
					else
						strsql = strsql & "username = '" & strSuspendName & "';"
				end if
				
				'create the objects used to suspend the account
				set rsSuspend = Server.CreateObject("ADODB.Recordset")
				
				'set the object attributes
				
				rsSuspend.LockType = 3
				rsSuspend.CursorType = 2
				
				'open the recordset
				rsSuspend.Open strsql, adocon
				
				'check to see if there is a record matching
				if rsSuspend.EOF then
						'there is no mathcig record
						Response.Redirect("admin_suspend_acc_form1.asp?error=5")
					else
						'send the email to say that the account has been suspended
						set objEmail = server.CreateObject("Jmail.SMTPMail")
						
						with objEmail
							.ServerAddress = "127.0.0.1"
							
							.Sender = "Support@worldsofwar.co.uk"
							.AddRecipient rsSuspend("email")
							.Subject = "Worlds of War II Account Suspended"
							.Body = "Your account for Worlds of War II with the username of " & rsSuspend("username") & " has been suspended by one of the administrators." & vbcrlf & vbcrlf & "The reason for your suspension was: " & vbcrlf & strReason & vbcrlf & vbcrlf & "If you have any questions, please ask in the #worldsofwar channel of our IRC network." & vbcrlf & vbcrlf & "Worlds of War Admins"
							.Execute 
						end with
						
						set objEmail = nothing
						
						'edit the users account so that their account has been suspended
						call rsSuspend.Update("suspended", cbool(true))
						
						Response.Redirect("admin.asp?msg=1")
				end if
				
				rsSuspend.Close
				
				set rssuspend = nothing
		end if
end if	

'code goes above here

end if
%>
<!---#include file="includes\end_check.asp"--->