<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
%>

<!---#include file="includes\top.asp"--->
<script language="javascript">
//<---
function checkForm(){
var rtnVar;
if (document.frmboard.contid.value==""){
	alert('You have not entered a continent to view.');
	rtnVal=false;
}else{
	rtnVal=true;
}

return rtnVal;
}
//--->
</script>

<center>
<BR />
<BR />

<table cellpadding="2" border="0" cellspacing="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-width:1; border-color;#FFFFFF;">
	<TR>
		<TD align="center" class="tdheaderline">
			View Continent Forum
		</TD>
	</TR>
	<TR>
		<TD class="norm" align="center">
			<form action="admin_view_board2.asp" name="frmboard" method="get" onsubmit="return checkForm();">
				What is the continent forum you would like to view?
				<BR />
				<input type="text" size="2" maxlength="4" name="contid" class="form" />
				<BR />
				<BR />
				<input type="submit" value="View Board" class="form" />
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