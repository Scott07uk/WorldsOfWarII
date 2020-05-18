<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
%>
<!---#include file="includes\top.asp"--->
<BR />
<BR />
<center>
	<table cellpadding="2" border="0" cellspacing="0" width="95%" style="border-top:solid; border-right:solid; border-left:solid; border-bottom:solid; border-width:1; border-color:#FFFFFF;">
		<TR>
			<TD class="tdheaderline" align="center">
				Edit Continent
			</TD>
		</TR>

		<TR>
			<TD class="norm" align="center">
				
				Please select the continent that you would like to edit.  
				<BR />
				<BR />
				<form action="admin_con_settings_form2.asp" method="post" name="frmSuspend">
					<table cellpadding="2" border="0" cellspacing="0">
						<TR>
							<TD class="tdheader" align="right">
								Continent Number
							</TD>
							<TD class="norm">
								<%
								Dim rscontinent
								
								strsql = "SELECT continentno, name FROM tblcontinents;"
								
								set rscontinent = server.CreateObject("ADODB.Recordset")
								
								rscontinent.Open strsql, adocon
								%>
								<SELECT name="continent" size="1" class="form">
								<%
									do while not rscontinent.EOF
								%>
									<option value="<% =rscontinent("continentno") %>">(<% =rscontinent("continentno") %>) <% =rscontinent("name") %></option>
								<%
									rscontinent.MoveNext 
								loop
								%>
								</select>
								<%
								rscontinent.Close 
								
								set rscontinent = nothing
								%>
							</TD>
						</TR>
					</table>
					<BR />
					<input type="submit" value="Edit Continent" class="form">
				</form>
				<BR />

				<BR />
				<A Href="admin.asp">Main Admin Menu</A>
			</TD>
		</TR>
	</table>
</center>

<!---#include file="includes\footer.asp"--->
<%
end if
%>
<!---#include file="includes\end_check.asp"--->