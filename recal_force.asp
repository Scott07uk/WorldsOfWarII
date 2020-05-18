<!---#include file="includes\check.asp"--->
<%
'SPELL CHECKED

Dim lngForceID
Dim rsunits
Dim rsaccount
Dim inttocontinent
Dim inttocountry
Dim strTargetName
Dim lngNewsID
Dim blnAccount

lngForceID = Request.QueryString("orderid")


if isnumeric(lngForceID) = false then
		Response.Redirect("mil_opps.asp")
	else
		set rsunits = server.CreateObject("ADODB.Recordset")
		
		rsunits.LockType = 3
		
		rsunits.CursorType = 2
		
		strsql = "SELECT * FROM tblmovements WHERE username = '" & strusername & "' AND orderno = " & lngforceid
		
		rsunits.Open strsql, adocon
		
		if rsunits.EOF then
				Response.Redirect("mil_opps.asp")
			else
				inttocontinent = cint(rsunits("tocontinent"))
				inttocountry = cint(rsunits("tocountry"))
				if cint(rsunits("eta")) = cint(rsunits("ticktime")) then
						set rsaccount = server.CreateObject("ADODB.Recordset")
						
						rsaccount.LockType = 3 
						
						rsaccount.CursorType = 2
						
						strsql = "SELECT numinfantry, numcommandos, numjeeps, nummat1, nummat2, numsherman, numspit1, numspit2, numspit5, numhurr, numhalifax, numwellington, numlancaster, numlandingcraft, numlandingship, numfrigates, numbattleship, numcarrier FROM tblaccounts WHERE username = '" & strusername & "';"
						
						rsaccount.Open strsql, adocon
						
						rsaccount.Fields("numinfantry") = clng(rsaccount("numinfantry")) + clng(rsunits("infantry"))
						rsaccount.Fields("numcommandos") = clng(rsaccount("numcommandos")) + clng(rsunits("comandoes"))
						rsaccount.Fields("numjeeps") = clng(rsaccount("numjeeps")) + clng(rsunits("jeeps"))
						rsaccount.Fields("nummat1") = clng(rsaccount("nummat1")) + clng(rsunits("matilda1"))
						rsaccount.Fields("nummat2") = clng(rsaccount("nummat2")) + clng(rsunits("matilda2"))
						rsaccount.Fields("numsherman") = clng(rsaccount("numsherman")) + clng(rsunits("sherman"))
						rsaccount.Fields("numspit1") = clng(rsaccount("numspit1")) + clng(rsunits("spitfire1"))
						rsaccount.Fields("numspit2") = clng(rsaccount("numspit2")) + clng(rsunits("spitfire2"))
						rsaccount.Fields("numspit5") = clng(rsaccount("numspit5")) + clng(rsunits("spitfire5"))
						rsaccount.Fields("numhurr") = clng(rsaccount("numhurr")) + clng(rsunits("hurricane"))
						rsaccount.Fields("numhalifax") = clng(rsaccount("numhalifax")) + clng(rsunits("halifax"))
						rsaccount.Fields("numwellington") = clng(rsaccount("numwellington")) + clng(rsunits("wellington"))
						rsaccount.Fields("numlancaster") = clng(rsaccount("numlancaster")) + clng(rsunits("lancaster"))
						rsaccount.Fields("numfrigates") = clng(rsaccount("numfrigates")) + clng(rsunits("frigates"))
						rsaccount.Fields("numbattleship") = clng(rsaccount("numbattleship")) + clng(rsunits("battleships"))
						rsaccount.Fields("numcarrier") = clng(rsaccount("numcarrier")) + clng(rsunits("carrier"))
						rsaccount.Fields("numlandingcraft") = clng(rsaccount("numlandingcraft")) + clng(rsunits("landingcraft"))
						rsaccount.Fields("numlandingship") = clng(rsaccount("numlandingship")) + clng(rsunits("landingship"))
						
						rsaccount.Update 
						
						rsunits.Delete 
						
						rsaccount.Close 
						
						set rsaccount = nothing
					else
						rsunits.Fields("status") = 3
						rsunits.Fields("ETA") = cint(cint(rsunits("ticktime")) - cint(rsunits("eta")))
						rsunits.Fields("tocountry") = lngcountryno
						rsunits.Fields("tocontinent") = lngcontinentno
				
						rsunits.Update
				end if
				
				rsunits.Close 
				
				strSQL = "SELECT username FROM tblaccounts WHERE continent = " & inttocontinent & " AND countryid = " & inttocountry
				
				rsunits.Open strsql, adocon
				
				if rsunits.EOF then
						blnAccount = false
					else
						blnAccount = true
						strtargetName = rsunits("username")
				end if
				
				rsunits.Close 
				
				strSQL = "SELECT * FROM tblnews ORDER BY newsID DESC;"
				
				rsunits.Open strsql, adocon
				
				if rsunits.EOF then
						lngnewsid = 1
					else
						lngnewsid = clng(rsunits("newsid")) + 1
				end if
				
				with rsunits
					.AddNew 
					.Fields("newsID") = lngnewsid
					.Fields("username") = strusername
					.Fields("timemade") = now()
					.Fields("message") = "Our force on its way to " & inttocontinent & ":" & inttocountry & " has been recalled.  "
					.Fields("unread") = 1
					
					.update
					
					if blnAccount = true then
							.AddNew 
							.Fields("newsID") = clng(lngnewsid + 1)
							.Fields("username") = strtargetname
							.Fields("timemade") = now()
							.Fields("message") = "An incomming force from " & lngcontinentno & ":" & lngcountryno & " has been recalled.  "
							.Fields("unread") = 1
							
							.Update 
					end if
				end with
				
				Response.Redirect("mil_opps.asp?error=7")
		end if
		
		rsunits.Close 
		
		set rsunits = nothing
end if	

%>
<!---#include file="includes\end_check.asp"--->