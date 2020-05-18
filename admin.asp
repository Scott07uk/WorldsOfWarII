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
<%
select case Request.QueryString("msg")
	case 1
		Response.Write("The account has been suspended and the owner emailed.  <BR /><BR />")
	case 2
		Response.Write("The account is no longer suspended.  <BR /><BR />")
	case 3
		Response.Write("The account settings have been updated.  <Br /><BR />")
	case 4
		Response.Write("The account has been relocated.  <BR /><BR />")
	case 5
		Response.Write("The account has been deleted.  <BR /><BR />")
	case 6
		Response.Write("The banner advert settings for the account have been updated.  <BR /><BR />")
	case 7
		Response.Write("The message has been deleted.  <BR /><BR />")
	case 8
		Response.Write("The selected messages have been deleted.  <BR /><BR />")
	case 9
		Response.Write("All continent topics and posts have been deleted.  <BR /><BR />")
	case 10
		Response.Write("The game settings have been updated.  <BR /><BR />")
	case 11
		Response.Write("The selected continent has been updated.  <BR /><BR />")
	case 12
		Response.Write("The message has been updated.  <BR /><BR />")
	case 13
		Response.Write("The game addresses have been updated.  <BR /><BR />")
	case 14
		Response.Write("The email address was banned.  <BR /><BR />")
	case 15
		Response.Write("The active email address database has been purged.  <BR /><BR />")
	case 16
		Response.Write("The server details were updated, please remember it can take up to 10 minutes for the changes to propagate the network.  <BR /><BR />")
	case 17
		Response.Write("The IRC server has been removed, please wait 10 minutes for the changes to propagate the network before you /rehash and remember to remove all leaf nodes this server has.<BR /><BR />")
	case 18
		Response.Write("The IRC server has been added, the server status will be determined once the changes propagate the network.  <BR /><BR />")
	case 19
		Response.Write("The banner was added to the rotation system.  <BR /><BR />")
	case 20
		Response.Write("The banner was removed from the rotation system.  <BR /><BR />")
	case 21
		Response.Write("The manual entry was updated.  <BR /><BR />")
	case 22
		Response.Write("The manual entry has been deleted from the manual.  <BR /><BR />")
	case 23
		Response.Write("The section was added to the manual.  <BR /><BR />")
	case 24
		Response.Write("The emails have been queued for dispatch.  <Br /><Br />")
	case 25
		Response.Write("The announcement has been posted.  <Br /><BR />")
		
end select
%>	
<Table cellpadding="2" cellspacing="0" border="0" width="95%">

	<TR>
		<TD class="tdheader" align="center" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			Account Options
		</TD>
		<TD class="tdheader" align="center" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			Message Options
		</TD>
	</TR>
		<TD valign="top" class="norm" style="border-bottom:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			<A HREF="admin_Suspend_acc_form1.asp">Suspend Account</A>
			<BR />
			<A HREF="admin_remove_suspention_form.asp">Remove Account Suspension</A>
			<BR />
			<A HREF="admin_edit_account_form1.asp">Edit Account Settings</A>
			<BR />
			<A HREF="admin_relocate_form1.asp">Relocate Account</A>
			<BR />
			<A HREF="admin_delete_account_form1.asp">Delete Account</A>
			<BR />
			<A HREF="admin_change_banner_ads.asp">Change Banner Advert Settings</A>
		</TD>
		<TD class="norm" valign="top" style="border-bottom:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			<A HREF="admin_view_all_msg1.asp">View All Messages</A>
			<BR />
			<A HREF="admin_del_old_msg_form1.asp">Delete Old Messages</A>
			<BR />
			<A HREF="admin_send_notice_form.asp">Send Admin Notice</a>
		</TD>
	</TR>
	<TR>
		<TD class="tdheader" align="center" valign="top" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			Board Options
		</TD>
		<TD class="tdheader" align="center" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			Game Options
		</TD>
	</TR>
	<TR>
		<TD class="norm" valign="top" style="border-bottom:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			<A HREF="admin_view_board1.asp">View Continent Board</A>
			<BR />
			<A HREF="admin_del_all_topics.asp" onclick="return confirm('Are you sure you want to delete ALL topics from ALL continents?');">Delete All Continent Topics</A>
		
		</TD>
		<TD class="norm" valign="top" style="border-bottom:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			<%
			Dim rsSettings		'object to hold the current settings
			
			set rsSettings = Server.CreateObject("ADODB.Recordset")
			
			'define the sql
			strSQL = "SELECT AllowedSignups, ticksEnabled, AttacksAllowed, DefenceAllowed, ResAllowed, ConAllowed FROM tblconfig;"
			
			rsSettings.Open strsql, adocon
			%>
			<A HREF="admin_change_setting.asp?type=ticks">
			<%
			if cbool(rssettings("ticksenabled")) = true then
					Response.Write("Stop")
				else	
					Response.write("Start")
			end if
			%> Ticks</A>
			<BR />
			<A HREF="admin_change_setting.asp?type=signup">
			<%
			if cbool(rssettings("allowedsignups")) = true then
					Response.Write("Stop")
				else
					Response.Write("Start")
			end if
			%>
			 Registration</A>
			<BR />
			<A HREF="admin_change_setting.asp?type=attacks">
			<%
			if cbool(rssettings("attacksallowed")) = true then
					Response.Write("Stop")
				else
					Response.Write("Start")
			end if
			%>
			Attacks</A>
			<BR />
			<A HREF="admin_change_setting.asp?type=defence">
			<%
			if cbool(rssettings("defenceallowed")) = true then
					Response.Write("Stop")
				else
					Response.Write("Start")
			end if
			%>
			Defence</A>
			<BR /> 
			<A HREF="admin_change_setting.asp?type=research">
			<%
			if cbool(rssettings("resallowed")) = true then
					Response.Write("Stop")
				else
					Response.Write("Start")
			end if
			%>
			Research</A>
			<BR />
			<A HREF="admin_change_setting.asp?type=construction">
			<%
			if cbool(rssettings("conallowed")) = true then
					Response.Write("Stop")
				else
					Response.Write("Start")
			end if
			%>
			Construction</A>
			
			<%
			rsSettings.Close 
			
			set rssettings = nothing
			%>
		</TD>
	</TR>
	<TR>
		<TD class="tdheader" align="center" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			Continent Options
		</TD>
		<TD class="tdheader" align="center" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			Texts
		</TD>
	</TR>
	<TR>
		<TD valign="top" class="norm" style="border-bottom:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			<A HREF="admin_con_settings_form1.asp">Continent Settings</A>
		</TD>
		<TD valign="top" class="norm" style="border-bottom:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			<A HREF="admin_change_msgs_form.asp?type=creator">Creator's Message</A>
			<Br />
			<A HREF="admin_change_msgs_form.asp?type=register">Registration Email</A>
			<BR />
			<A HREF="admin_game_urls_form.asp">Game URLs</A>
		</TD>
	</TR>
	<TR>
		<TD class="tdheader" align="center" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			Security
		</TD>
		<TD class="tdheader" align="center" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			IRC
		</TD>
	</TR>
	<TR>
		<TD valign="top" class="norm" style="border-bottom:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			<A HREF="admin_ban_email.asp">Ban Email Address</A>
			<BR />
			<A HREF="admin_del_active_email.asp" onclick="return confirm('Are you sure you want to delete ALL the activated email addresses?');">Delete Active Email</A>
		</TD>
		<TD valign="top" class="norm" style="border-bottom:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			<A HREF="admin_irc_list.asp">Manage Server List</A>
			<Br />
			<A HREF="admin_add_irc_form.asp">Add IRC Server</A>
		</TD>
	</TR>
	<TR>
		<TD class="tdheader" align="center" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			Banner Adverts
		</TD>
		<TD class="tdheader" align="center" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			Manual
		</TD>
	</TR>
	<TR>
		<TD valign="top" class="norm" style="border-bottom:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			<A HREF="admin_add_banner_form.asp">Add Banner Advert</A>
			<BR />
			<A HREF="admin_list_banners.asp">List All Banner Adverts</A>
		</TD>
		<TD valign="top" class="norm" style="border-bottom:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			<A HREF="admin_change_manual_form.asp">Change Manual Entry</A>
			<BR />
			<A HREF="admin_add_manual_form.asp">Add Manual Entry</A>
		</TD>
	</TR>
	<TR>
		<TD class="tdheader" align="center" style="border-top:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			Other
		</TD>
	</TR>
	<TR>
		<TD valign="top" class="norm" style="border-bottom:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
			<A HREF="admin_read_reports.asp?type=abuse">Read Abuse Reports</A>
			<BR />
			<A HREF="admin_read_reports.asp?type=complaint">Read Complaints</A>
			<BR />
			<A HREF="admin_read_reports.asp?type=suggestion">Read Suggestions</A>
			<Br />
			<A HREF="admin_post_announcement.asp">Post Announcement</A>
			<BR />
			<A HREF="admin_email_all_users.asp">Email All Users</a>
		</TD>
	</TR>

</table>


<!---#include file="includes\footer.asp"--->

<%
end if
%>
<!---#include file="includes\end_check.asp"--->