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
			Add New Banner Advert
		</TD>
	</TR>
	<TR>
		<TD class="norm" align="center">
			Fill in the form below to add a new banner advert to the site.  
			<br />
			<form action="admin_add_banner_engine.asp" method="post">
				<table cellpadding="2" border="0" cellspacing="0">
					<TR>
						<TD class="norm" align="right">
							Banner HTML Code
						</TD>
						<TD class="norm">
							<TEXTAREA cols="50" rows="5" name="banner" class="form"></TEXTAREA>
						</TD>
					</TR>
				</table>
				<BR />
				<input type="submit" value="Update" class="form" / id=submit1 name=submit1>
			</form>
			
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