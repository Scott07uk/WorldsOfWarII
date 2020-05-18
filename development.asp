<!---#include file="includes/check.asp"--->
<!---#include file="includes/top.asp"--->

<!---SPELL CHECKED--->
<center>
<table width="95%" cellpadding="0" border="0" cellspacing="0" style="border-top:none; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD align="center" class="norm">
			<A HREF="javascript:openHelp('9');">Help With This Page</A>
		</TD>
	</TR>
</table>
</center>
<BR />
<form action="startdevelopmentengine.asp" method="post">
<center>
<br />
<%
select case Request.QueryString("error")
	case 1
		Response.Write("You did not order any industry.  <BR /><BR />")
	case 2
		Response.Write("You do not have enough land to start construction.  <BR /><BR />")
	case 3
		Response.Write("You do not have enough food to start construction.  <BR /><BR />")
	case 4
		Response.Write("You do not have enough wood to start construction.  <BR /><BR />")
	case 5
		Response.Write("You do not have enough iron to start construction.  <BR /><BR />")
	case 6
		Response.Write("You do not have enough money to start construction.  <BR /><BR />")
	case 7
		Response.Write("The industry has been ordered for construction.  <BR /><BR />")
	case 8
		Response.Write("You must entered a number for the amount you wish to convert. <BR /><BR />")
	case 9
		Response.Write("You can only exchange resources that you already have.  <BR /><BR />")
	case 10
		Response.Write("The resources were successfully traded.  <BR /><BR />")
	case 11
		response.write("You can not trade one resource for the same resource.  <BR /><BR />")
	case 12
		Response.write("You can not trade resources for a minus amount.  <BR /><BR />")
	case 13
		Response.Write("The tick is currently processing, you can not order more units or change resources during this time, please try again in a few seconds.  <BR /><BR />")
		
end select
Dim rsBuilding

set rsBuilding = server.CreateObject("ADODB.Recordset")

strsql = "SELECT farms, sawmills, mines, refineries FROM tblpendingbuildings WHERE username = '" & strusername & "';"

Dim lngLandUsed

rsBuilding.Open strsql, adocon

lngLandUsed = 0

do while not rsBuilding.EOF 
	lngLandUsed = lngLandUsed + clng(rsbuilding("farms")) + clng(rsbuilding("sawmills")) + clng(rsbuilding("mines")) + clng(rsbuilding("refineries"))
	
	rsBuilding.Movenext 
loop

rsBuilding.Close 
	
	strsql = "SELECT numfarms, nummines, numrefineries, nummills FROM tblaccounts WHERE username = '" & strusername & "';"
	
	rsBuilding.Open strsql, adocon
	
	lngLandUsed = lngLandused + clng(rsbuilding("numfarms")) + clng(rsbuilding("nummines")) + clng(rsbuilding("numrefineries")) + clng(rsbuilding("nummills"))
	%>

<table align="center" bgcolor="#0D1724" cellpadding="2" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<tr>
		<td class="tdheaderline" align="center" colspan="8">
			Order Industry
		</td>
	</tr>
	<tr>
		<td colspan="8" align="center" class="all">
			Each industry requires 1 land.
			You have a total of <% =formatnumber(lngland, 0) %> land.
			You have a total of <% =formatnumber(lnglandused,0) %> industry.
			You have <% =formatnumber(clng(lngland-lnglandused),0) %> land free for new industry.
		</td>
	</tr>
	<tr>
		<td class="all" width="30%" valign="top">
			Industry
		</td>
		<td class="all" width="10%" valign="top">
			Food
		</td>
		<td class="all" width="10%" valign="top">
			Wood
		</td>
		<td class="all" width="10%" valign="top">
			Iron
		</td>
		<td class="all" width="10%" valign="top">
			Money
		</td>
		<td class="all" width="10%" valign="top">
			Ticks
		</td>
		<td class="all" width="13%" valign="top">
			Currently Owned
		</td>
		<td class="all" width="7%" valign="top">&nbsp;
			
		</td>
	</tr>
	<tr>
		<td class="norm">
			Farm
		</td>
		<td class="norm">
			250
		</td>
		<td class="norm">
			100
		</td>
		<td class="norm">
			100
		</td>
		<td class="norm">
			100
		</td>
		<td class="norm">
			8 ticks
		</td>
		<td class="norm">
			<%=rsbuilding("numFarms")%>
		</td>
		<td class="norm">
			<input type="text" size="5" value="0" name="farms" class="form">
		</td>
	</tr>
	<tr>
		<td class="norm">
			Sawmill
		</td>
		<td class="norm">
			100
		</td>
		<td class="norm">
			0
		</td>
		<td class="norm">
			250
		</td>
		<td class="norm">
			500
		</td>
		<td class="norm">
			10 ticks
		</td>
		<td class="norm">
			<%=rsbuilding("numMills")%>
		</td>
		<td class="norm">
			<input type="text" size="5" value="0" name="sawmills" class="form">
		</td>
	</tr>
	<tr>
		<td class="norm">
			Refinery
		</td>
		<td class="norm">
			100
		</td>
		<td class="norm">
			0
		</td>
		<td class="norm">
			500
		</td>
		<td class="norm">
			750
		</td>
		<td class="norm">
			10 ticks
		</td>
		<td class="norm">
			<%=rsbuilding("numrefineries")%>
		</td>
		<td class="norm">
			<input type="text" size="5" value="0" name="Refineries" class="form">
		</td>
	</tr>
	<tr>
		<td class="norm">
			Mine
		</td>
		<td class="norm">
			500
		</td>
		<td class="norm">
			50
		</td>
		<td class="norm">
			150
		</td>
		<td class="norm">
			200
		</td>
		<td class="norm">
			6 ticks
		</td>
		<td class="norm">
			<%=rsbuilding("numMines")%>
		</td>
		<td class="norm">
			<input type="text" size="5" value="0" name="Mines" class="form">
		</td>
	</tr>

	<tr>
		<td colspan="8" style="border-top: solid; border-width: 1px;" align="center">
			<input type="submit" value="Start" class="form">
		</td>
	</tr>
</table>

</form>
<BR />
<SCRIPT language="javascript">
//<!---begin hiding

function checkExchange() {
var rtnVal;

if (document.frmExch.from.value==document.frmExch.to.value){
	rtnVal=false;
	alert('You can not trade one resourece for the same resource.');
}else{
	if (document.frmExch.amount.value==""){
		rtnVal=false;
		alert('You have not entered an ammount to change.');
	}else{
		rtnVal = true;
	}
}

return rtnVal;
}

//---->
</SCRIPT>
<Table align="center" cellpadding="2" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-right:solid; border-left:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline" align="center">
			Trade Resources:
		</TD>
	</TR>
	<TR>
		<TD class="norm" align="center">
			This allows you to trade your resources for other resources for a fee of 25%.
			<BR />
			<form action="trade_engine.asp" method="post" name="frmExch" onsubmit="return checkExchange();">
				Trade
				<select name="from" size="1" class="form">
					<option value="1">Food</option>
					<option value="2">Wood</option>
					<option value="3">Iron</option>
					<option value="4">Money</option>
				</SELECT>
				For
				<select name="to" size="1" class="form">
					<option value="1">Food</option>
					<option value="2">Wood</option>
					<option value="3">Iron</option>
					<option value="4">Money</option>
				</SELECT>
				Amount to Change
				<input type="text" size="5" maxlength="8" name="amount" class="form" value="0" />
				<input type="submit" value="Exchange" class="form" />
			</form>
		</TD>
	</TR>
</Table>
<BR />
<BR />
<table align="center" bgcolor="#0D1724" cellpadding="2" cellspacing="0" width="95%" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<tr>
		<td class="tdheaderline" align="center" colspan="7">
			Industry Ordered
		</td>
	</tr>
	<TR>
		<TD class="norm" align="center">
			Below is a table of all the industry currently under construction.
				<BR />
				<table cellpadding="2" cellspacing="0" border="0">
					<TR>
						<TD class="tdheaderline">
							Industry
						</TD>
						<%
						Dim intCounter
						intCounter =1 
						
						for intCounter = 1 to 10
							Response.Write("<TD class=""tdheaderline"" align=""center"" width=""17px"">")
							Response.Write(intCounter)
							Response.Write("</TD>")
						next
						%>
						
					</TR>
					<% 
					rsBuilding.Close 
					strsql = "SELECT ticksleft, farms, sawmills, mines, refineries FROM tblpendingbuildings WHERE username = '" & strusername & "' ORDER by ticksleft ASC;"
					
					rsBuilding.open strsql, adocon
					
					Dim strUnitArrival(4,11)
					Dim intProcessTimes(4)
					Dim yCounter
					Dim xCounter
					
					yCounter = 1
					xCounter = 1
					for xcounter = 0 to 3
						for ycounter = 0 to 10
							strUnitArrival(xcounter,ycounter) = "&nbsp;"
						next
					next
					strUnitArrival(0,0) = "Farms"
					strUnitArrival(1,0) = "Saw Mills"
					strUnitArrival(2,0) = "Refineries"
					strUnitArrival(3,0) = "Mines"
					
					
					intProcessTimes(0) = 8
					intProcessTimes(1) = 10
					intProcessTimes(2) = 10
					intProcessTimes(3) = 6
					
					for ycounter = 1 to 11
						if not rsBuilding.EOF then
								if rsBuilding("ticksleft") = ycounter then
										if clng(rsBuilding("farms")) <> 0 then
												strunitArrival(0,ycounter) = rsBuilding("farms")
										end if
										if clng(rsBuilding("sawmills")) <> 0 then
												strunitArrival(1,ycounter) = rsBuilding("sawmills")
										end if
									
										if clng(rsBuilding("refineries")) <> 0 then
												strunitArrival(2,ycounter) = rsBuilding("refineries")
										end if
										if clng(rsBuilding("mines")) <> 0 then
												strunitArrival(3,ycounter) = rsBuilding("mines")
										end if
																				
										rsBuilding.MoveNext
								end if
						end if
					next
					%>
					<TR>
					<%
					ycounter = 0
					xcounter = 0
					for xcounter = 0 to 3
						Response.Write("<TR>")
						for ycounter = 0 to 10
							if intProcessTimes(xcounter) >= ycounter then
									Response.Write("<TD class=form align=center>")
									Response.Write(strunitarrival(xcounter,ycounter))
									Response.Write("</TD>")
								else
							end if

						next
						Response.Write("</TR>")
					next
					%>
					</TR>
				</table>
			</TD>
		</TR>
	</table>
</center>
<%

rsbuilding.Close
set rsbuilding = nothing
%>
<!---#include file="includes/footer.asp"--->
<!---#include file="includes/end_check.asp"--->
