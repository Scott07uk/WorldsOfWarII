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
<script language="javascript">
//<---
function checkForm(){

var rtnVar;

if (document.frmirc.numeric.value==""){rtnVar=false;alert('You have not entered a server numeric for this server. ');
}else{if (document.frmirc.name.value==""){rtnVar=false;alert('You have not entered a network address on which this server can be accessed');
}else{if (document.frmirc.adminname.value==""){rtnVar=false;alert('You have not entered the name of the person responsible for the upkeep of this server.');
}else{if (document.frmirc.adminemail.value==""){rtnVar=false;alert('You have not entered an email address for '+document.frmirc.adminname.value);
}else{rtnVar=confirm('Are you sure this server is linked and ready to run on the network?');
}
}
}
}
return rtnVar;
}
//--->
</script>
<%
select case Request.QueryString("error")
	case 1
		Response.Write("You did not enter all the information, ALL the fields are required.  <BR /><BR />")
	case 2
		Response.Write("<font size=""4"" color=""#FF0000""><B><U>!!!A SERVER WITH THAT NUMERIC ALREADY EXISTS, DISCONNECT YOUR SERVER IMMEDIATELY!!!</U></B></FONT><BR /><BR />")
end select
%>
<table cellpadding="2" width="95%" border="0" cellspacing="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline" align="center">
			Add New IRC Server
		</TD>
	</TR>
	<TR>
		<TD class="norm" align="center">
			Fill in all the fields below to add a new IRC server to the network. Remember to allow the server to link BEFORE you add it here.  
			<br />
			<b>REMEMBER ALL CHANGES CAN TAKE UP TO 10 MINUTES TO PROPAGATE THE NETWORK.</b>
			<br />
			<form action="admin_add_irc_engine.asp" method="post" name="frmirc" onsubmit="return checkForm();">
				<table cellpadding="2" border="0" cellspacing="0">
					<TR>
						<TD class="norm" align="right">
							Server Numeric
						</TD>
						<TD class="norm">
							<input type="text" size="3" maxlength="3" name="numeric" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right">
							Network Name
						</TD>
						<TD class="norm">
							<input type="text" size="40" maxlength="200" name="name" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right">
							Server Admin
						</TD>
						<TD class="norm">
							<input type="text" size="15" maxlength="50" name="adminname" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right">
							Contact Email
						</TD>
						<TD class="norm">
							<input type="text" size="25" maxlength="150" name="adminemail" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right">
							Allow Connections?
						</TD>
						<TD class="norm">
							<SELECT name="connectallow" size="1" class="form">
								<option value="True" SELECTED>Yes</option>
								<option value="False">No</option>
							</SELECT>
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right">
							Provides Services?
						</TD>
						<TD class="norm">
							<SELECT name="services" size="1" class="form">
								<option value="True">Yes</option>
								<option value="False" SELECTED>No</option>
							</SELECT>
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right">
							Provides Statistics?
						</TD>
						<TD class="norm">
							<SELECT name="stats" size="1" class="form">
								<option value="True">Yes</option>
								<option value="False" SELECTED>No</option>
							</SELECT>
						</TD>
					</TR>
					<TR>
						<TD class="norm" align="right">
							Status?
						</TD>
						<TD class="norm">
							<SELECT name="status" size="1" class="form">
								<option value="2">Offline</option>
								<option value="3" SELECTED>De-Configured</option>
							</SELECT>
						</TD>
					</TR>
				</table>
				<BR />
				<input type="submit" value="Update" class="form" />
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