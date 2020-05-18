<!---#include file="includes/check.asp"--->

<!---SPELL CHECKED--->
<%
Dim lngShowNum
Dim rscontinent 'Get continent banner image
Set rscontinent = Server.CreateObject("ADODB.Recordset")
StrSQL = "SELECT Banner, Name, ContinentNo FROM TblContinents WHERE ContinentNo = "
If Request.Form("continent") <> "" Then
	If IsNumeric(Request.Form("continent")) = False Then
		lngShowNum = lngContinentNo
	Else
		lngShowNum = Request.Form("continent")
	End If
Else
	lngShowNum = lngContinentNo
End If
strSQL = strSQL & lngShowNum
rscontinent.Open StrSQL, Adocon

If rscontinent.EOF Then
	Response.Redirect("continent.asp")
Else
%>

<!---#include file="includes/top.asp"--->
<center>
<table width="95%" cellpadding="0" border="0" cellspacing="0" style="border-top:none; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD align="center" class="norm">
			<A HREF="javascript:openHelp('17');">Help With This Page</A>
		</TD>
	</TR>
</table>
</center>
<br>
<center>

Select Continent
<br>

<form action="continent.asp" method="post" name="changecontinent">
<input type="text" name="continent" size="1" maxlength="4" class="form" value="<% = lngShowNum %>">
<input type="submit" value="Go" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1; padding:0;" class="form">
</form>

<table cellpadding="2" border="0" width="95%" cellspacing="0" style="border-top:solid; border-left:solid; border-right:solid; border-bottom:solid; border-color:#FFFFFF; border-width:1;">
<tr>
<td class="tdheaderline" align="center">
<% = rscontinent("Name") %>
</td>
</tr>

<tr>
<td class="norm" align="center">
<img src="<% =rscontinent("Banner") %>" alt="<% = rscontinent("Name") %>">
</td>
</tr>
</table>
<%
rscontinent.Close
%>

<table cellpadding="2" border="0" width="95%" cellspacing="0" style="border-top:solid; border-left:solid; border-right:solid; border-bottom:solid; border-color:#FFFFFF; border-width:1;">
<tr>
<td class="tdheaderline" width="10%">
Co-ords
</td>

<td class="tdheaderline" width="30%">
Leader
</td>

<td class="tdheaderline" width="30%">
Country
</td>

<td class="tdheaderline" width="10%">
Score
</td>

<td class="tdheaderline" width="10%">
Land
</td>
</tr>

<br>

<%
StrSQL = "SELECT Continent, CountryID, LeaderName, CountryName, Score, AmountLand, PowerBase FROM TblAccounts WHERE Continent = " & lngShowNum
rscontinent.Open StrSQL, Adocon
%>

<%
Do While Not rscontinent.EOF
%>
<tr>
<td class="norm" width="10%">
<%
If rscontinent("PowerBase") = True Then
	Response.Write("<font color=#ff0000>")
End If
%>
<% = rscontinent("Continent") %>:<% = rscontinent("CountryID") %>
<%
If rscontinent("PowerBase") = True Then
	Response.Write("</font>")
End If
%>
</td>

<td class="norm" width="30%">
<%
If rscontinent("PowerBase") = True Then
	Response.Write("<font color=#ff0000>")
End If
%>
<% = rscontinent("LeaderName") %>
<%
If rscontinent("PowerBase") = True Then
	Response.Write("</font>")
End If
%>
</td>

<td class="norm" width="30%">
<%
If rscontinent("PowerBase") = True Then
	Response.Write("<font color=#ff0000>")
End If
%>
<% = rscontinent("CountryName") %>
<%
If rscontinent("PowerBase") = True Then
	Response.Write("</font>")
End If
%>
</td>

<td class="norm" width="10%">
<%
If rscontinent("PowerBase") = True Then
	Response.Write("<font color=#ff0000>")
End If
%>
<% =formatnumber(rscontinent("score"),0) %>
<%
If rscontinent("PowerBase") = True Then
	Response.Write("</font>")
End If
%>
</td>

<td class="norm" width="10%">
<%
If rscontinent("PowerBase") = True Then
	Response.Write("<font color=#ff0000>")
End If
%>
<% =formatnumber(rscontinent("AmountLand"),0) %>
<%
If rscontinent("PowerBase") = True Then
	Response.Write("</font>")
End If
%>
</td>
</tr>
<%
rscontinent.MoveNext
Loop
%>

</table>
<br>
<font color="#ff0000">Red = Power Base</font>
</center>
<!---#include file="includes/footer.asp"--->
<%
End If
rscontinent.Close
Set rscontinent = Nothing
%>
<!---#include file="includes/end_check.asp"--->