<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else

dim rsmanual
Dim lngid

lngid =Request.QueryString("ID")

if isnumeric(lngid) = false then
		Response.Redirect("admin_edit_manual_form.asp")
	else
	
	strsql = "SELECT entryname, entry, ordering, onsidebar FROM tblmanual WHERE enteryid = " & lngid
	
	set rsmanual = server.CreateObject("ADODB.Recordset")
	
	rsmanual.Open strsql, adocon
	
	if rsmanual.EOF then
			Response.Redirect("admin_edit_manual_form.asp")
		else
		
%>
<!---#include file="includes\top.asp"--->


<BR />
<BR />
<center>

<table cellpadding="2" width="95%" border="0" cellspacing="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline" align="center">
			Edit Manual
		</TD>
	</TR>
	<TR>
		<TD class="norm" align="center">
			Here you can make changes to the manual entries.  
			<BR />
			<form action="admin_edit_manual_engine.asp" method="post">
				<input type="hidden" name="id" value="<% =lngid %>" />
				<table cellpadding="2" border="0" cellspacing="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;" width="95%">
					<TR>
						<TD class="norm" align="right">
							Entry Title
						</TD>
						<TD class="norm">
							<input type="text" size="20" maxlength="50" name="title" class="form" value="<% =rsmanual("entryname") %>" />
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right">
							Order
						</TD>
						<TD class="norm">
							<input type="text" size="2" name="order" maxlength="3" class="form" value="<% =rsmanual("ordering") %>" />
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right">
							On Sidebar?
						</TD>
						<TD class="norm">
							<SELECT name="onsidebar" size="1" class="form">
								<% if cbool(rsmanual("onsidebar")) = true then %>
									<option value="True" selected>Yes</option>
									<option value="False">No</option>
								<% else %>
									<option value="True">Yes</option>
									<option value="False" selected>No</option>
								<% end if %>
							</select>
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right" valign="top">
							Manual Entry
						</TD>
						<TD class="norm">
							<textarea name="entry" cols="70" class="form" rows="11"><% =rsmanual("entry") %></TEXTAREA>
						</TD>
					</TR>
				</table>
				<BR />
				<BR />
				<input type="submit" value="Update" class="form" />
			</form>
			<A Href="admin.asp">Main Admin Menu</A>
			
		</TR>
	</TR>
</table>
</center>

<!---#include file="includes\footer.asp"--->

<%
end if
rsmanual.Close 

set rsmanual = nothing
end if
end if
%>
<!---#include file="includes\end_check.asp"--->