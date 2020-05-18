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

<table cellpadding="2" width="95%" border="0" cellspacing="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline" align="center">
			Manage Banners
		</TD>
	</TR>
	<TR>
		<TD class="norm" align="center">
			You can use this page to view and delete all the banners that are active on the system. 
			<BR />
			<table cellpadding="2" border="0" cellspacing="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;" width="95%">
				<TR>
					<TD class="tdheaderline">
						ID
					</TD>
					<TD class="tdheaderline">
						Banner
					</TD>
					<TD class="tdheaderline">
						Added By
					</TD>
					<TD class="tdheaderline">
						Delete
					</TD>
				</TR>
				<%
				Dim rsbanners
				
				set rsbanners = server.CreateObject("ADODB.Recordset")
				
				strsql = "SELECT adid, adtext, addedby FROM tblbanners;"
				
				rsbanners.Open strsql, adocon
				
				do while not rsbanners.EOF 
				%>
				<TR>
					<TD class="norm">
						<% =rsbanners("adid") %>
					</TD>
					<TD class="norm">
						<% =rsbanners("adtext") %>
					</TD>
					<TD class="norm">
						<% =rsbanners("addedby") %>
					</TD>
					<TD class="norm">
						<A HREF="admin_del_banner_engine.asp?id=<% =rsbanners("adid") %>">Delete</A>
					</TD>
				</TR>
				<%
				rsbanners.MoveNext 
				loop
				
				rsbanners.Close
				set rsbanners = nothing
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