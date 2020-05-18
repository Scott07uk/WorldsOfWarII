<!---#include file="includes/check.asp"--->
<!---#include file="includes/top.asp"--->

<!---SPELL CHECKED--->

<BR />
<form action="newtopicengine.asp" method="post">

<table align="center" bgcolor="#0D1724" cellpadding="2" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<tr>
		<td class="tdheader" colspan="2">
			New Topic
		</td>
	</tr>
<% 
if request.querystring("err") = 1 then 
%>
	<tr>
		<td class="all" colspan="2" align="center">
			<center><font color="#ff0000">The subject field is empty.</font></center>
		</td>
	</tr>
<% 
end if
if request.querystring("err") = 2 then 
%>
	<tr>
		<td class="all" colspan="2" align="center">
			<center><font color="#ff0000">The message field is empty.</font></center>
		</td>
	</tr>
<%
end if
%>
	<tr>
		<td class="all" style="border-right: solid; border-width: 1;">
			Subject
		</td>
		<td class="all">
			<input type="text" name="subject" size="50" class="form" valign="top" />
		</td>
	</tr>
	<tr>
		<td class="all" style="border-right: solid; border-width: 1;" valign="top">
			Message
		</td>
		<td class="all">
			<textarea name="message" rows="5" cols="70" class="form"></textarea>
		</td>
	</tr>
	<tr>
		<td colspan="2" align="center" width="100%">
			<input type="submit" value="Create thread" class="form" />
			<input type="reset" value="Reset fields" class="form" />
		</td>
	</tr>
</table>
<!---#include file="includes/footer.asp"--->
<!---#include file="includes/end_check.asp"--->
