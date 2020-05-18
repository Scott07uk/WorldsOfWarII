<!---#include file="includes/check.asp"--->
<!---#include file="includes/top.asp"--->

<!---SPELL CHECKED--->
<center>
<table width="95%" cellpadding="0" border="0" cellspacing="0" style="border-top:none; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD align="center" class="norm">
			<A HREF="javascript:openHelp('8');">Help With This Page</A>
		</TD>
	</TR>
</table>
</center>
<%
dim rsExplore
Dim lngExplorers

strsql = "SELECT numExplorers from tblaccounts where username = '" & strusername & "';"

'create the objects
set rsexplore = server.CreateObject("ADODB.Recordset")

rsExplore.Open strsql, adocon

lngExplorers = clng(rsexplore("numexplorers"))

rsExplore.Close

%>

<script language="javascript">
//<--being hiding

function checkOrder(){
//function to make sure the ordering new explorers form is correct
var rtnVar;
if (document.explore.explorers.value==""){
	rtnVar=false;
	alert('You have not entered how many units you would like to order.');
}
else{
	rtnVar=true;
}

return rtnVar;
}

function calCost(){
var initialCost=0;
var curLand=<% =lngland %>;
var landWanted = document.exploreland.acres.value;
var cost=0;

initialCost=curLand;
cost=initialCost;

if (landWanted<1){
document.exploreland.display.value = "You cannot explore less than 1 acre."
}
else{

for (i=1; i<landWanted; i++){
	cost=cost+initialCost+i;
}
document.exploreland.display.value="Exploring "+document.exploreland.acres.value+" acres will require "+cost+" explorers.";
}
return true;
}


//-->

</script>
<BR />
<BR />
<center>
<%
select case Request.QueryString("error")
	case 1
		Response.Write("You did not enter a number of units to train. <BR /><BR />")
	case 2
		Response.Write("The number of units to train was invalid. <BR /><BR />")
	case 3
		Response.Write("You do not have enough money to train that many units. <BR /><BR />")
	case 4
		Response.Write("You do not have enough food to train that many units.  <BR /><BR />")
	case 5
		Response.Write("You do not have enough wood to train that many units. <BR /><BR />")
	case 6
		Response.Write("You do not have enough iron to train that many units. <BR /><BR />")
	case 7
		Response.Write("The number of units you wanted have been enrolled for training. <BR /><BR />")
	case 8
		Response.Write("You did not enter an amount of land to explore.  <BR /><BR />")
	case 9
		Response.Write("You amount of land to explore was invalid.  <BR /><BR />")
	case 10
		Response.Write("You do not have enough explorers to explore that much land.  <BR /><BR />")
	case 11
		if Request.QueryString("gold") = 0 then
				Response.Write("Your explorers have explored and taken control of the land you requested. <BR /><BR />")
			else
				Response.Write("Your explorers have explored and taken control of the land you requested. <BR>In the process they discovered " & Request.QueryString("gold") & " gold! <Br /><BR />")
		end if
	case 12
		Response.Write("The tick is current processing so you can not order new units, please wait a few seconds and try again. <BR /><BR />")
		
end select

%>
	<table cellpadding="2" width="95%" cellspacing="0" border="0" style="border-top:solid; border-bottom:solid; border-right:solid; border-left:solid; border-color:#FFFFFF; border-width:1;">
		<TR>
			<TD class="tdheaderline" align="center">
				Training Explorers
			</TD>
		</TR>
		<TR>
			<TD class="norm" align="center">
				<%
				'open the pending units table to see how many explorers there are
				strsql = "SELECT explorers, timeleft from tblpendingunits WHERE username = '" & strusername & "' AND Explorers > 0 ORDER BY timeleft ASC;"
				
				rsExplore.Open strsql, adocon
				
				if rsExplore.EOF then
				%>
				You have no explorers currently in training.
				<%
				else
				do while not rsExplore.EOF 
				if clng(rsexplore("explorers")) > 0 then
				%>
					You currently have <% =formatnumber(rsexplore("explorers"),0) %> in training.<BR> Training will be completed in <% =rsexplore("timeleft") %> ticks.  
					<BR />
				<%
				end if
				rsExplore.MoveNext 
				loop
				
				end if
				
				rsExplore.Close 
				
				set rsexplore = nothing
				%>
			</TD>
		</TR>
	</table>
	<BR />
	<BR />
	<table width="95%" cellpadding="2" cellspacing="0" border="0" style="border-top:solid; border-bottom:solid; border-right:solid; border-left:solid; border-color:#FFFFFF; border-width:1;">
		<TR>
			<TD class="tdheaderline" align="center">
				Train New Explorers
			</TD>
		</TR>
		<TR>
			<TD class="norm" align="center">
				<form action="order_explore.asp" method="post" name="explore" onsubmit="return checkOrder();">
					<table cellpadding="5" border="0" width="100%" cellspacing="0">
						<TR>
							<TD class="tdheaderline">
								Unit Name
							</TD>
							<TD class="tdheaderline">
								Food
							</TD>
							<TD class="tdheaderline">
								Wood
							</TD>
							<TD class="tdheaderline">
								Iron
							</TD>
							<TD class="tdheaderline">
								Money
							</TD>
							<TD class="tdheaderline">
								Time
							</TD>
							<TD class="tdheaderline">
								Order
							</TD>
						</TR>
						<TR title="Explorers allow you to explore and take control of more land, allowing you to build more industry.">
							<TD class="norm">
								Explorers
							</TD>
							<TD class="norm">
								250
							</TD>
							<TD class="norm">
								10
							</TD>
							<TD class="norm">
								50
							</TD>
							<TD class="norm">
								500
							</TD>
							<TD class="norm">
								6 ticks
							</TD>
							<TD class="norm">
								<input type="text" size="2" name="explorers" class="form" maxlength="3" />
							</TD>
						</TR>
					</table>
					<BR />
					<input type="submit" value="Order" class="form">
				</form>
			</TD>
		</TR>
	</table>
	<BR />
	<BR />
	<table width="95%" cellpadding="2" cellspacing="0" border="0" style="border-top:solid; border-bottom:solid; border-right:solid; border-left:solid; border-color:#FFFFFF; border-width:1;">
		<TR>
			<TD class="tdheaderline" align="center">
				Explore New Land
			</TD>
		</TR>
		<TR>
			<TD class="norm" align="center">
				The greater the amount of land you own, the more explorers you will need to bring more under your control.
				<BR />
				You currently have <% =lngexplorers %> you can use to explore new land.  
				<Br />
				<Br />
				<form action="explore_land.asp" method="post" name="exploreland">
					Number of acres to explore:
					<input type="text" size="2" class="form" name="acres" maxlength="3" value="0" />
					<BR />
					<BR />
					<input type="button" class="form" name="Cal" value="Calculate" onclick="calCost();" />
					<input type="text" size="60" name="display" class="form" readonly />
					<BR />
					<BR />
					<input type="submit" value="Explore" class="form" />
				</form>
			</TD>
		</TR>
	</table>
</center>
<BR />

<!---#include file="includes/footer.asp"--->			
<!---#include file="includes/end_check.asp"--->