<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
%>

<!---#include file="includes\top.asp"--->
<center>
<BR />
<BR />
<table cellpadding="2" border="0" width="95%" cellspacing="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
	<TR>
		<TD class="tdheaderline" align="center">
			Post Announcements
		</TD>
	</TR>
	<TR>
		<TD class="norm" align="center">
			<form action="post_announcement_engine.asp" method="post">
				Use this form to post an announcement on the portal. 
				<BR />
				<table cellpadding="2" cellspacing="0" border="0">
					<TR>
						<TD class="tdheader" align="right">
							Subject
						</TD>
						<TD class="norm">
							<input type="text" size="50" name="subject" maxlength="150" class="form" value="[WoWII]:" />
						</TD>
					</TR>
					<TR>
						<TD class="tdheader" align="right" valign="top">
							Message
						</TD>
						<TD class="norm">
							<textarea name="message" cols="70" rows="10" class="form"></TEXTAREA>
						</TD>
					</TR>
				</table>
				<BR />
				<input type="submit" value="Send Message" class="form" />
			</form>
		</TD>
	</TR>
</table>
</center>


<!---#include file="includes\footer.asp"--->

<%
end if
%>
<!---#include file="includes\end_check.asp"--->