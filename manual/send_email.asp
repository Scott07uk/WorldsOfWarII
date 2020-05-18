<%
if (trim(Request.Form("email")) = "") OR (trim(Request.Form("subject")) = "") OR (trim(Request.Form("details")) = "") then
		Response.Redirect("contact.asp")
	else
		Dim objEmail
		
		set objEmail = server.CreateObject("JMail.SMTPMAIL")
		
		with objEmail
			.ServerAddress ="127.0.0.1"
			.Sender = Request.Form("email")
			.AddRecipient "support@worldsofwar.co.uk"
			.Subject = Request.Form("subject")
			.Body = server.HTMLEncode(Request.Form("details"))
			
			.Execute 
		end with
		
		set objemail = nothing
		%>
		
		<!---#include file="inc_manual_top.asp"--->
<table cellpadding="2" border="0" cellspacing="0" width="95%" style="border-left:solid; border-right:solid; border-top:solid; border-bottom:solid; border-width:1; border-color:#FFFFFF;">
	<TR>
		<TD class="tdheaderline" align="center">
			Worlds of War II Manual
		</TD>
	</TR>
	<TR>
		<TD class="norm">
				You message has been sent, thank you. 
		</TD>
	</TR>
</table>
<!---#include file="inc_manual_bottom.asp"--->
<%
end if
%>