<!---#include file="includes\check.asp"--->
<%
'<!---SPELL CHECKED--->
if cbool(blnadmin) = false then

		Response.Redirect("over.asp")
else
%>
<html>
	<head>
		<title></title>
		<link rel=stylesheet HREF="includes/all.css">
	</head>
<Body bgcolor="black" link="white" alink="white" vlink="white" text="#ffffff">
<%
dim rsmanual 

strsql = "SELECT username, comment FROM tblmanualcomments WHERE manid = " & Request.QueryString("ID")

set rsmanual = server.CreateObject("ADODB.Recordset")

rsmanual.Open strsql, adocon

if rsmanual.EOF then
%>
	<Center>
		<font face="arial" size="2">
			There are no comments for this part of the manual.
		</font>
		<BR />
		<BR />
<%
else
%>
<center>
<%
do while not rsmanual.EOF 
%>
<table cellpadding="2" cellspacing="0" border="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-width:1; border-color:#FFFFFF;">
	<TR>
		<TD class="tdheaderline">
			Comment Left By
			<%
			if rsmanual("username") = "" then
					Response.Write("Unknown")
				else
					Response.Write(rsmanual("username"))
			end if
			%>
		</TD>
	</TR>
	<TR>
		<TD class="norm">
			<% =rsmanual("comment") %>
		</TD>
	</TR>
</table>
<BR />
<BR />
<%
rsmanual.MoveNext 
loop
%>
<A HREF="admin_del_manual_comments.asp?id=<% =Request.QueryString("id") %>">Delete Comments</A>
<%
end if

rsmanual.Close 

set rsmanual = nothing

%>
<A HREF="javascript:window.close();">Close Window</a>
</center>

</body>
</HTML>
<%
end if
%>
<!---#include file="includes\end_check.asp"--->