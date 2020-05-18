<!---#include file="includes/check.asp"--->
<!---#include file="includes/top.asp"--->

<!---SPELL CHECKED--->
<center>
<table width="95%" cellpadding="0" border="0" cellspacing="0" style="border-top:none; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD align="center" class="norm">
			<A HREF="javascript:openHelp('18');">Help With This Page</A>
		</TD>
	</TR>
</table>
</center>
<br />
<center>
	Top 10 Most Powerful Nations
</center>
<BR />
<table align="center" cellpadding="2" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<tr>
		<td class="tdheader" width="10%" style="border-right: solid; border-width:1;">
			Rank
		</td>
		<td class="tdheader" width="10%" style="border-right: solid; border-width:1;">
			Co-ordinates
		</td>
		<td class="tdheader" width="30%" style="border-right: solid; border-width:1;">
			Country
		</td>
		<td class="tdheader" width="30%" style="border-right: solid; border-width:1;">
			Leader
		</td>
		<TD class="tdheader" width="10%" style=" border-right:solid; border-width:1;">
			Land
		</TD>
		<td class="tdheader" width="10%" style="border-width:1;">
			Score
		</td>
		
	</tr>
<%
Dim intCounter 		'varible used to loop through 10 records
set rsworld = server.createobject("ADODB.recordset")

strSql = "SELECT tblaccounts.leadername, countryname, countryID, continent, score, amountland FROM tblAccounts ORDER BY score DESC;"

rsworld.open StrSQL, adocon

for counter = 1 to 10
	if not rsworld.eof then
%>
	<tr>
		<td class="tdheader" width="0%" style="border-top: solid; border-right: solid; border-width:1;">
			<% =counter %>
		</td>
		<td class="norm" style="border-top: solid; border-right: solid; border-width:1;">
			[<%=rsworld("continent")%>:<%=rsworld("CountryID")%>]
		</td>
		<td class="norm" style="border-top: solid; border-right: solid; border-width:1;">
			<%=rsworld("countryname")%>
		</td>
		<td class="norm" style="border-top: solid; border-right: solid; border-width:1;">
			<%=rsworld("leadername")%>
		</td>
		<TD class="norm" style="border-top:solid; border-right:solid; border-width:1;">
			<% =rsworld("amountland") %>
		</TD>
		<td class="norm" style="border-top: solid; border-width:1;">
			<%=formatnumber(rsworld("score"),0)%>
		</td>
		
	</tr>
<%
	rsworld.MoveNext
end if
next

rsworld.close
%>
</table>
<BR />
<BR />

<center>
	Top 10 Biggest Nations
</center>
<BR />
<table align="center" cellpadding="2" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<tr>
		<td class="tdheader" width="10%" style="border-right: solid; border-width:1;">
			Rank
		</td>
		<td class="tdheader" width="10%" style="border-right: solid; border-width:1;">
			Co-ordinates
		</td>
		<td class="tdheader" width="30%" style="border-right: solid; border-width:1;">
			Country
		</td>
		<td class="tdheader" width="30%" style="border-right: solid; border-width:1;">
			Leader
		</td>
		<TD class="tdheader" width="10%" style=" border-right:solid; border-width:1;">
			Land
		</TD>
		<td class="tdheader" width="10%" style="border-width:1;">
			Gold
		</td>
		
	</tr>
<%
strSql = "SELECT tblaccounts.leadername, countryname, countryID, continent, gold, amountland FROM tblAccounts ORDER BY amountland DESC;"

rsworld.open StrSQL, adocon

for counter = 1 to 10
	if not rsworld.eof then
%>
	<tr>
		<td class="tdheader" width="0%" style="border-top: solid; border-right: solid; border-width:1;">
			<% =counter %>
		</td>
		<td class="norm" style="border-top: solid; border-right: solid; border-width:1;">
			[<%=rsworld("continent")%>:<%=rsworld("CountryID")%>]
		</td>
		<td class="norm" style="border-top: solid; border-right: solid; border-width:1;">
			<%=rsworld("countryname")%>
		</td>
		<td class="norm" style="border-top: solid; border-right: solid; border-width:1;">
			<%=rsworld("leadername")%>
		</td>
		<TD class="norm" style="border-top:solid; border-right:solid; border-width:1;">
			<% =rsworld("amountland") %>
		</TD>
		<td class="norm" style="border-top: solid; border-width:1;">
			<%=rsworld("Gold")%>
		</td>
		
	</tr>
<%
	rsworld.MoveNext
end if
next

rsworld.close
set rsworld = nothing
%>
</table>
<BR />
<BR />
<!---#include file="includes/footer.asp"--->
<!---#include file="includes/end_check.asp"--->