<!---#include file="includes/check.asp"--->

<!---SPELL CHECKED--->

<%

Dim rsmessages 'Delete messages.
Dim strBox

strBox = Request.Form("box")
If strBox <> "" Then
	Set rsmessages = Server.CreateObject("ADODB.Recordset")
	strBox = Replace(strBox, ",", " OR MessageID =")
	
	StrSQL = "SELECT TblMessages.* FROM TblMessages WHERE MessageID = " & strBox
	rsmessages.LockType = 3
	rsmessages.CursorType = 2
	rsmessages.Open StrSQL, Adocon
	
	Do While Not rsmessages.EOF
		If rsmessages("ToContinent") = LngContinentNo Then
			If rsmessages("ToCountry") = LngCountryNo Then
				rsmessages.Delete
				rsmessages.MoveNext
			End If
		End If
	Loop
	
	rsmessages.Close
	Set rsmessages = Nothing
	Response.Redirect("msgs.asp?msg=1")
Else
	Response.Redirect("msgs.asp")
End If
%>

<!---#include file="includes/end_check.asp"--->