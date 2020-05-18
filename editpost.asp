<!---#include file="includes/check.asp"--->
<!---#include file="includes/top.asp"--->

<!---SPELL CHECKED--->

<%
'DECLAIR YOUR SODDING VARIBLES
'DONT USE THE BLOODY *
Dim rsTopic

Set rstopic = Server.CreateObject("ADODB.Recordset")

StrSQL = "SELECT tblconpost.message FROM tblconpost WHERE TopicID = " & request.Querystring("topicID") & "AND PostID = " & request.querystring("msgID")

rstopic.open strSQL, adocon

If not rstopic.EOF then
strMsg = Replace(rstopic("message"),"<BR>",vbCrlf)
%>

<br />
<form action="Editengine.asp?topicid=<%=request.querystring("topicID")%>&msgID=<%=request.querystring("msgID")%>" method="post">
<table align="center" bgcolor="#0D1724" cellpadding="3" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<tr>
		<td class="tdheaderline">
			Edit Post
		</td>
	</tr>
	<tr>
		<td class="all">
			<textarea name="message" rows="7" cols="70" class="form"><%=StrMsg%></textarea>
		</td>
	</tr>
	<tr>
		<td align="center">
			<input type="submit" value="Edit" class="form">
		</td>
	</tr>
</table>
</form>

<%
end if 
rstopic.close
set rstopic = nothing
%>
<!---#include file="includes/footer.asp"--->
<!---#include file="includes/end_check.asp"--->