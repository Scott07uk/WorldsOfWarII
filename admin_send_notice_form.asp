<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
%>	

<!---#include file="includes\top.asp"--->
<br />
<center>
<SCRIPT language="javascript">

function checkForm(){
var rtnVal;

if (document.frmmsg.continent.value==""){
	rtnVal=false;
	alert('You have not entered a location to send the message to.  ');
}else{
	if (document.frmmsg.subject.value==""){
		rtnVal=false;
		alert('You have not entered a subject for the message.  ');
	}else{
		if (document.frmmsg.body.value==""){
			rtnVal=false;
			alert('You have not entered a message to send to the user.');
		}else{
			rtnVal=true;
		}
	}
}
return rtnVal;
}

</script>
<BR />
<BR />
<% select case request.querystring("error")
	case 1
		Response.Write("You did not complete all the required message fields.  <BR /><BR />")
	case 2
		Response.Write("The continent does not exist.  <BR /><BR />")
	case 3
		Response.Write("You did not enter a valid country number.  <BR /><BR />")
end select
%>
<form action="admin_send_msg_engine.asp" method="post" name="frmmsg" onsubmit="return checkForm();">
	<table style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;" cellpadding="5" cellspacing="5" width="75%">
		<tr>
			<td align="right" class="tdheader">
				Co-ords
			</td>
			<td class="norm">
				<input type="text" name="continent" size="1" maxlength="4" class="form">:<input type="text" name="country" size="1" maxlength="2" class="form">
			</td>
		</tr>
		<tr>
			<td align="right" class="tdheader">
				Subject
			</td>
			<td class="norm">
				<input type="text" name="subject" size="65" maxlength="250" class="form">
			</td>
		</tr>
		<tr>
			<td align="right" valign="top" class="tdheader">
				Message
			</td>
			<td class="norm">
				<textarea name="body" cols="60" rows="8" class="form"></textarea>
			</td>
		</tr>
		<tr>
			<td>
			</td>
			<td class="norm">
				<input type="submit" value="Send" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1; padding:0;" class="form">&nbsp;&nbsp;
				<input type="reset" value="Reset" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1; padding:0;" class="form">
			</td>
		</tr>
	</table>
</form>


<!---#include file="includes\footer.asp"--->

<%
end if
%>
<!---#include file="includes\end_check.asp"--->