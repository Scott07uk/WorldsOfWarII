<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else

'code goes below here

'to do list
'delete all the messages to the account			DONE
'delete the attack results						DONE
'delete all the intelegance reports				DONE
'delete the account news						DONE
'delete pending buildings						DONE
'delete pending units							DONE
'delete the votes								DONE
'delete the account							

'Declair the varibles
Dim rsDel
Dim strDelName
Dim lngContinent
Dim lngCountry
Dim strVoteName

'get the username to delete
strDelName = Request.QueryString("username")

if strDelName = "" then
		Response.Redirect("admin_delete_account_form1.asp")
	else
		'generate the sql script to get the account details
		StrSQL = "SELECT continent, countryID, vote FROM tblaccounts WHERE username = '" & strdelname & "';"
		
		'create the objects to delete the account
		
		set RsDel = Server.CreateObject("ADODB.Recordset")
		
		'set the object attributes
		rsDel.LockType = 3
		rsDel.CursorType = 2
		
		rsDel.Open strsql, adocon
		
		if rsDel.EOF then
				Response.Redirect("admin_delete_account_form1.asp?error=1")
			else
				lngContinent = rsDel("continent")
				lngCountry = rsDel("Countryid")
				strVoteName = rsdel("vote")
				
				rsDel.Close
				
				'select all the messages to the account
				strsql = "SELECT * FROM tblmessages WHERE tocontinent = " & lngcontinent & " AND tocountry = " & lngcountry
				
				rsDel.Open strsql, adocon
				
				do while not rsDel.EOF
				
					rsDel.Delete
					
					rsDel.MoveNext 
				loop
				
				rsDel.Close 
				
				
				'clear out the attacks results
				strSQL = "SELECT tblattackresults.* FROM tblattackresults WHERE username = '" & strdelname & "';"
				
				rsDel.Open strsql, adocon
				
				do while not rsDel.eof 
					rsDel.Delete 
					rsDel.Movenext
				loop
				
				rsDel.close
				
				strsql = "SELECT TblIntelegenceReports.* FROM TblIntelegenceReports WHERE username = '" & strdelname & "';"
				
				rsDel.Open strsql, adocon
				
				'delete all the records
				
				do while not rsDel.EOF 
					rsDel.Delete
					rsDel.MoveNext 
				loop
				
				rsDel.Close 
				
				strSql = "SELECT tblnews.* FROM tblnews WHERE username = '" & strdelname & "';"
				
				rsDel.Open strsql, adocon
				
					'delete all the records
				do while not rsDel.EOF 
					rsDel.Delete
					rsDel.Movenext
				loop
				
				rsDel.Close 
				
				strSql = "SELECT tblpendingunits.* FROM tblpendingunits WHERE username = '" & strdelname & "';"
				
				rsDel.Open strsql, adocon
				
				'delete all the records
				do while not rsDel.EOF
					rsDel.Delete
					rsDel.Movenext
				loop
				
				rsDel.Close
				
				'delete the pending buildings
				
				strsql = "SELECT tblpendingbuildings.* FROM tblpendingbuildings WHERE username = '" & strdelname & "';"
				
				rsDel.Open strsql, adocon
				
				'delete all the records
				do while not rsDel.EOF 
					rsDel.Delete
					rsDel.MoveNext 
				loop
				
				rsDel.Close 
				
				strsql = "SELECT numvotes FROM tblaccounts WHERE username = '" & strVoteName & "';"
				
				rsDel.Open strsql, adocon
				
				if rsDel.EOF then
						'do nothing
					else
						if rsdel("numvotes") > 0 then
								call rsDel.Update("numvotes", cint(rsdel("numvotes")) - 1)
						end if
				end if
				
				rsDel.Close 
				
				strSQL = "SELECT tblaccounts.* FROM tblaccounts WHERE username = '" & strdelname & "';"
				
				rsDel.Open strsql, adocon
				
				rsDel.Delete 
				
				rsDel.Close
				
				Response.Redirect("admin.asp?msg=5")
		end if
				
		rsDel.Close
	
		set rsdel = nothing
end if				
				
		

'code goes above here

end if
%>
<!---#include file="includes\end_check.asp"--->