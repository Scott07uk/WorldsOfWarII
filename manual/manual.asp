<!---#include file="inc_manual_top.asp"--->
<SCRIPT language="javascript">
	//<--hide script from older browsers
			
	//function to open a new window
			
	function openWin(url)
	{
		//open the new window
		window.open(url,'manrate', config='height=450,width=650,toolbar=no');
	}
			
			
	//--end hiding>
</script>
<%
Dim strErrMessage

strErrMessage = "Sorry but the link you clicked on to, to get this page is not valid, please select a valid entry from whats shown on the sidebar."
if isnumeric(Request.QueryString("view")) = false then

Response.Write(strErrMessage)

else

strsql = "SELECT entry FROM tblmanual WHERE enteryid = " & Request.QueryString("view")

rsmanual.open strsql, adocon

if rsmanual.eof then
		Response.Write(strErrMessage)
	else
		if (isnull(rsManual("entry")) = true) OR (trim(rsManual("entry")) = "") then
				Response.Write("Sorry but this manual entry has not been written, please check again later.  ")
			else
				Response.Write(rsManual("entry"))
		
				Response.Write("<BR /><BR />")
		
				Response.Write("<table cellpadding=""2"" border=""0"" cellspacing=""0"" width=""95%"" style=""border-top:solid; border-bottom:solid; border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;""><TR><TD class=""norm"" align=""right"">")
				Response.Write("Please help us improve the manual by rating this section on a scale of 0-5: &nbsp; ")
				Response.Write("<A HREF=""javascript:openWin('rate_form.asp?id=" & Request.QueryString("view") & "&value=0');"">0</A>&nbsp;")
				Response.Write("<A HREF=""javascript:openWin('rate_form.asp?id=" & Request.QueryString("view") & "&value=1');"">1</A>&nbsp;")
				Response.Write("<A HREF=""javascript:openWin('rate_form.asp?id=" & Request.QueryString("view") & "&value=2');"">2</A>&nbsp;")
				Response.Write("<A HREF=""javascript:openWin('rate_form.asp?id=" & Request.QueryString("view") & "&value=3');"">3</A>&nbsp;")
				Response.Write("<A HREF=""javascript:openWin('rate_form.asp?id=" & Request.QueryString("view") & "&value=4');"">4</A>&nbsp;")
				Response.Write("<A HREF=""javascript:openWin('rate_form.asp?id=" & Request.QueryString("view") & "&value=5');"">5</A>&nbsp;")
				Response.Write("</TD>")
				Response.Write("</TR></table>")
		end if
end if

rsmanual.close 

end if
%>
<!---#include file="inc_manual_bottom.asp"--->