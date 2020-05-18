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
			New Manual Page
		</TD>
	</TR>
	<TR>
		<TD class="norm" align="center">
			Here you can make the needed changes to the manual entries.  
			<BR />
			<form action="admin_add_manual_engine.asp" method="post">
				<table cellpadding="2" border="0" cellspacing="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;" width="95%">
					<TR>
						<TD class="norm" align="right">
							Entry Title
						</TD>
						<TD class="norm">
							<input type="text" size="20" maxlength="50" name="title" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right">
							Order
						</TD>
						<TD class="norm">
							<input type="text" size="2" name="order" maxlength="3" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right">
							On Sidebar?
						</TD>
						<TD class="norm">
							<SELECT name="onsidebar" size="1" class="form">
									<option value="True" selected>Yes</option>
									<option value="False">No</option>
							</select>
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right" valign="top">
							Manual Entry
						</TD>
						<TD class="norm">
							<textarea name="entry" cols="70" class="form" rows="11"></TEXTAREA>
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
%>
<!---#include file="includes\end_check.asp"--->