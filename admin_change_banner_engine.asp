<!---#include file="includes/check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else

'code goes below here

'declair varibles
Dim strChange			'username to change
Dim rsChange			'recordset to perform the change
Dim blnSetting			'the setting the user is to be set to

'get the user input
strChange = Request.Form("username")
blnSetting = Request.Form("setting")

'format the input
blnSetting = cbool(blnsetting)
strchange = trim(strChange)

if strChange = "" then
		Response.Redirect("admin_change_banner_form.asp?error=1")
	else
		'create the objects to handle the editing
		set rsChange = server.CreateObject("ADODB.Recordset")
		
		'set the object attributes to allow updating
		rsChange.LockType = 3
		
		rsChange.CursorType = 2
		
		'generate the sql script
		strSQL = "SELECT showads FROM tblaccounts WHERE username = '" & strchange & "';"
		
		'open the collection
		rsChange.Open strSQL, adocon
		
		'check to make sure the username is valid
		if rsChange.EOF then
				Response.Redirect("admin_change_banner_form.asp?error=2")
			else
				'perfrom the edit
				call rsChange.Update("showads", blnSetting)
				Response.Redirect("admin.asp?msg=6")
		end if
		
		'kill the object
		rsChange.Close
		set rsChange = nothing
end if


'code goes above here

end if
%>

<!---#include file="includes/end_check.asp"--->

