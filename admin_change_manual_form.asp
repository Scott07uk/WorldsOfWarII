<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
%>
<!---#include file="includes\top.asp"--->

<SCRIPT language="javascript">
	//&lt;--hide script from older browsers
			
	//function to open a new window
			
	function openWin(url)
	{
		//open the new window
		window.open(url,'email', config='height=450,width=650,toolbar=no');
	}
						
	//--end hiding&gt;
</script>

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
			Select the section of the manual you want to edit or delete.
			<BR />
			<table cellpadding="2" border="0" cellspacing="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;" width="95%">
				<TR>
					<TD class="tdheaderline">
						ID
					</TD>
					<TD class="tdheaderline">
						Section Title
					</TD>
					<TD class="tdheaderline">
						On Sidebar?
					</TD>
					<TD class="tdheaderline">
						Rating
					</TD>
					<TD class="tdheaderline">
						Comments
					</TD>
					<TD class="tdheaderline">
						Edit?
					</TD>
					<TD class="tdheaderline">
						Delete?
					</TD>
				</TR>
				<%
				Dim rsmanual
				
				set rsmanual = server.CreateObject("ADODB.Recordset")
				
				strsql = "SELECT enteryid, entryname, totalrating, numrating, onsidebar FROM tblmanual;"
				
				rsmanual.Open strsql, adocon
				
				do while not rsmanual.EOF 
				%>
				<TR>
					<TD class="norm">
						<% =rsmanual("enteryid") %>
					</TD>
					<TD class="norm">
						<% =rsmanual("entryname") %>
					</TD>
					<TD class="norm">
						<% 
						if cbool(rsmanual("onsidebar")) = true then
								Response.Write("Yes")
							else
								Response.Write("No")
						end if
						%>
					<TD class="norm">
						<% 
						if rsmanual("numrating") = 0 then
								Response.Write("0.0000")
							else
								Response.Write(formatnumber(cdbl(clng(rsmanual("totalrating")) / clng(rsmanual("numrating"))),4))
						end if
						%>
					</TD>
					<TD class="norm">
						<A HREF="javascript:openWin('admin_view_manual_comments.asp?id=<% =rsmanual("enteryid") %>');">Comments</A>
					</TD>
					<TD class="norm">
						<A HREF="admin_edit_manual_form.asp?id=<% =rsmanual("enteryid") %>">Edit</A>
					</TD>
					<TD class="norm">
						<A HREF="admin_delete_manual_engine.asp?id=<% =rsmanual("enteryid") %>" onclick="return confirm('Are you sure you want to delete this entry on the manual?');">Delete</A>
					</TD>
					
				</TR>
				<%
				rsmanual.MoveNext 
				loop
				
				rsmanual.Close
				set rsmanual = nothing
				%>
			</table>
			<BR />
			<BR />
			<A Href="admin.asp">Main Admin Menu</A>
			
		</TR>
	</TR>
</table>
</center>

<!---#include file="includes\footer.asp"--->

<%
end if
%>
<!---#include file="includes\end_check.asp"--->