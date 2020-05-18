
<!---#include file="includes\check.asp"--->
<!---#include file="includes\top.asp"--->

<!---SPELL CHECKED--->
<center>
<table width="95%" cellpadding="0" border="0" cellspacing="0" style="border-top:none; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD align="center" class="norm">
			<A HREF="javascript:openHelp('14');">Help With This Page</A>
		</TD>
	</TR>
</table>
</center>
<%
Dim rsunits
strsql = "SELECT numinfantry, numcommandos, numjeeps, nummat1, nummat2, numsherman, numspit1, numspit2, numspit5, numhurr, numhalifax, numwellington, numlancaster, numlandingcraft, numlandingship, numfrigates, numbattleship, numcarrier FROM tblaccounts WHERE username = '" & strusername & "';"

set rsunits = server.CreateObject("ADODB.Recordset")

rsunits.Open strsql, adocon

%>
<Br />
<BR />
<center>
<%
select case Request.QueryString("error")
	case 1
		Response.Write("You have not entered a valid target nation.<BR /><BR />")
	case 2
		Response.Write("You did not allocate your task force any units.<BR /><BR />")
	case 3 
		Response.Write("The target nation does not exist.<BR /><BR />")
	case 4
		Response.Write("The selected target is too small for you to attack.<BR /><BR />")
	case 5
		Response.Write("You do not have enough landing craft or landing ships to launch an attack of that size. You need " & cstr(Request.QueryString("ship")) & " landing ships and " & cstr(Request.QueryString("craft")) & " landing craft.<BR /><BR />")
	case 6
		Response.write("The task force has been launched.<BR /><BR />")
	case 7
		Response.Write("The task force has been recalled.<BR /><BR />")
	case 8
		Response.Write("There are too many people launching an attack at that location for you to send forces.<BR /><BR />")
	case 9
		Response.Write("The tick engine is currently processing, you can not launch a new task force during this time, please try again in a few seconds.<BR /><BR />")
		
end select
%>
<Script language="javascript">
//<---
function checkForm() {
var rtnVar;

if (document.frmmission.country.value=="") {
	rtnVar=false;
	alert('You have not entered a valid country to attack.');
}else{
	if (document.frmmission.continent.value=="") {
		rtnVar=false;
		alert('You have not entered a valid country to attack. ');
	}else{
		if (document.frmmission.continent.value==<% =lngcontinentno %>){
			if (document.frmmission.country.value==<% =lngcountryno %>){
				rtnVar=false;
				alert('You cannot launch an attack against yourself.');
			}else{
				rtnVar=true;
			}
		}else{
			rtnVar=true;
		}
	}
}
return rtnVar;

}
//--->
</script>
<form action="new_mission.asp" method="post" onsubmit="return checkForm();" name="frmmission">
	<table align="center" cellpadding="2" cellspacing="0" width="95%" border="0" style="border-top:solid; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
		<TR>
			<TD class="tdheaderline" align="center">
				New Mission
			</TD>
		</TR>
		<TR>
			<TD class="norm" align="center">
				<table cellpadding="2" border="0" cellspacing="0">
					<TR>
						<TD class="tdheaderline">
							Unit Name
						</TD>
						<TD class="tdheaderline">
							Description
						</TD>
						<TD class="tdheaderline">
							Local ETA
						</TD>
						<TD class="tdheaderline">
							Global ETA
						</TD>
						<TD class="tdheaderline">
							Home Stock
						</TD>
						<TD class="tdheaderline">
							Number to Send
						</TD>
					</TR>
					<TR>
						<TD class="norm">
							Infantry
						</TD>
						<TD	class="norm">
							<font size="1">
							<I>
							Infantry are the smallest unit and most common unit on the battlefield.
							</I>
							</font>
						</TD>
						<TD class="norm">
							5
						</TD>
						<TD class="norm">
							Requires Landing Craft (8)
						</TD>
						<TD class="norm">
							<% =formatnumber(clng(rsunits("numinfantry")),0) %>
						</TD>
						<TD class="norm">
							<input type="text" size="7" maxlength="8" name="infantry" value="0" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm">
							Commandos
						</TD>
						<TD class="norm">
							<font size="1">
							<I>
								Commandos are a better trained and armed ground unit.
							</I>
							</font>
						</TD>
						<TD class="norm">
							5
						</TD>
						<TD class="norm">
							Requires Landing Craft (8)
						</TD>
						<TD class="norm">
							<% =formatnumber(clng(rsunits("numcommandos")),0) %>
						</TD>
						<TD class="norm">
							<input type="text" size="7" maxlength="8" name="commandoes" value="0" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm">
							Jeeps:
						</TD>
						<TD class="norm">
							<font size="1">
							<I>
								The jeep is the cheapest motorised unit. It is highly mobile and has mounted guns.
							</I>
							</font>
						</TD>
						<TD class="norm">
							3
						</TD>
						<TD class="norm">
							Requires Landing Ship (10)
						</TD>
						<TD class="norm">
							<% =formatnumber(clng(rsunits("numjeeps")),0) %>
						</TD>
						<TD class="norm">
							<input type="text" size="7" maxlength="8" name="jeeps" value="0" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm">
							Matilda MkI
						</TD>
						<TD class="norm">
							<font size="1">
							<I>
								The Matilda MkI is a slow moving excessively well armoured mobile machine gun.
							</I>
							</font>
						</TD>
						<TD class="norm">
							4
						</TD>
						<TD class="norm">
							Requires Landing Ship (10)
						</TD>
						<TD class="norm">
							<% =formatnumber(clng(rsunits("nummat1")),0) %>
						</TD>
						<TD class="norm">
							<input type="text" size="7" maxlength="8" name="mat1" value="0" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm">
							Matilda MkII
						</TD>
						<TD class="norm">
							<font size="1">
							<I>
								The Matilda MkII is a mobile fortress with a sound 2lb turret mounted gun.
							</I>
							</font>
						</TD>
						<TD class="norm">
							4
						</TD>
						<Td class="norm">
							Requires Landing Ship (10)
						</TD>
						<TD class="norm">
							<% =formatnumber(clng(rsunits("nummat2")),0) %>
						</TD>
						<TD class="norm">
							<input type="text" size="7" maxlength="8" name="mat2" value="0" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm">
							Sherman Tank
						</TD>
						<TD class="norm">
							<font size="1">
								<I>
									The Sherman Tank faster and more powerful than the Matilda, but lacks armour.
								</I>
							</font>
						</TD>
						<TD class="norm">
							3
						</TD>
						<TD class="norm">
							Requires Landing Ship (10)
						</TD>
						<TD class="norm">
							<% =formatnumber(clng(rsunits("numsherman")),0) %>
						</TD>
						<TD class="norm">
							<input type="text" size="7" maxlength="8" name="sherman" value="0" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm">
							Spitfire MkI
						</TD>
						<TD class="norm">
							<font size="1">
								<I>
									The Spitfire MkI is the first high speed, all metal single wing fighter planes.
								</I>
							</font>
						</TD>
						<TD class="norm">
							2
						</TD>
						<TD class="norm">
							4
						</TD>
						<TD class="norm">
							<% =formatnumber(clng(rsunits("numspit1")),0) %>
						</TD>
						<TD class="norm">
							<input type="text" size="7" maxlength="8" name="spitfire1" value="0" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm">
							Spitfire MkII
						</TD>
						<TD class="norm">
							<font size="1">
								<I>
									The MkII Spitfire now features air to air cannons with improved flight performance.
								</I>
							</font>
						</TD>
						<TD class="norm">
							2
						</TD>
						<TD class="norm">
							4
						</TD>
						<TD class="norm">
							<% =rsunits("numspit2") %>
						</TD>
						<TD class="norm">
							<input type="text" size="7" maxlength="8" name="spitfire2" value="0" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm">
							Spitfire MkV
						</TD>
						<TD class="norm">
							<font size="1">
								<I>
									The MkV Spitfire now incorporates under wing rockets for strafing ground targets.
								</I>
							</font>
						</TD>
						<TD class="norm">
							2
						</TD>
						<TD class="norm">
							4
						</TD>
						<TD class="norm">
							<% =formatnumber(clng(rsunits("numspit5")),0) %>
						</TD>
						<TD class="norm">
							<input type="text" size="7" maxlength="8" name="spitfire5" value="0" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm">
							Hurricane
						</TD>
						<TD class="norm">
							<font size="1">
								<I>
									The Hawker Hurricane is a new type of single wing plane which is very durable and quick to repair.
								</I>
							</font>
						</TD>
						<TD class="norm">
							2
						</TD>
						<TD class="norm">
							4
						</TD>
						<TD class="norm">
							<% =formatnumber(clng(rsunits("numhurr")), 0) %>
						</TD>
						<TD class="norm">
							<input type="text" size="7" maxlength="8" name="hurricane" value="0" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm">
							Halifax
						</TD>
						<TD class="norm">
							<font size="1">
								<I>
									The Halifax bomber is a slow, cumbersome long range bomber.
								</I>
							</font>
						</TD>
						<TD class="norm">
							3
						</TD>
						<TD class="norm">
							6
						</TD>
						<TD class="norm">
							<% =rsunits("numhalifax") %>
						</TD>			
						<Td class="norm">
							<input type="text" size="7" maxlength="8" name="halifax" value="0" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm">
							Wellington
						</TD>
						<TD class="norm">
							<font size="1">
								<I>
									The Wellington bomber is a more modern, adaptable medium range bomber.
								</I>
							</font>
						</TD>
						<TD class="norm">
							3
						</TD>
						<TD class="norm">
							6
						</TD>
						<TD class="norm">
							<% =formatnumber(clng(rsunits("numwellington")),0) %>
						</TD>
						<TD class="norm">
							<input type="text" size="7" maxlength="8" name="wellington" value="0" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm">
							Lancaster
						</TD>
						<TD class="norm">
							<font size="1">
								<I>
									The Lancaster is a tough, long range bomber which uses the G navigation system for extra accuracy.
								</I>
							</font>
						</TD>
						<TD class="norm">
							3
						</TD>
						<TD class="norm">
							6
						</TD>
						<TD class="norm">
							<% =formatnumber(clng(rsunits("numlancaster")),0) %>
						</TD>
						<TD class="norm">
							<input type="text" size="7" maxlength="8" name="lancaster" value="0" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm">
							Frigates
						</TD>
						<TD class="norm">
							<font size="1">
								<I>
									The Frigates are smaller war ships often used as escorts to assist the main ships.
								</I>
							</font>
						</TD>
						<TD class="norm">
							6
						</TD>
						<TD class="norm">
							9
						</TD>
						<TD class="norm">
							<% =formatnumber(clng(rsunits("numfrigates")),0) %>
						</TD>
						<TD class="norm">
							<input type="text" size="7" maxlength="8" name="frigates" value="0" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm">
							Battleships
						</TD>
						<TD class="norm">
							<font size="1">
								<I>
									The pride of the navy, the large battleships, designed for sinking its own class and decimating coastlines.
								</I>
							</font>
						</TD>
						<TD class="norm">
							6
						</TD>
						<TD class="norm">
							9
						</TD>
						<TD class="norm">
							<% =formatnumber(clng(rsunits("numbattleship")),0) %>
						</TD>
						<TD class="norm">
							<input type="text" size="7" maxlength="8" name="battleships" value="0" class="form" />
						</TD>
					</TR>
					<TR>
						<TD class="norm">
							Carrier
						</TD>
						<TD class="norm">
							<font size="1">
								<I>
									The largest ships in the navy, used to reshape costlines and bomb any target well out of the range of any gun.
								</I>
							</font>
						</TD>
						<TD class="norm">
							7
						</TD>
						<TD class="norm">
							12
						</TD>
						<TD class="norm">
							<% =formatnumber(clng(rsunits("numcarrier")),0) %>
						</TD>
						<Td class="norm">
							<input type="text" size="7" maxlength="8" name="carrier" value="0" class="form" />
						</TD>
					</TR>
				</table>
				<BR />
				The time the units take to arrive will be the speed of the slowest unit.<BR>Aerial attacks can arrive quickly with little time for preparation, while sea attacks take longer and give the enemy more warning.
				<BR />
				Mission Type
				<select name="type" class="form" size="1">
					<option value="0" selected>---</option>
					<option value="1">Attack</option>
					<option value="2">Defend</option>
				</select>
				<BR />
				Target
				<input type="text" name="continent" size="3" maxlength="4" class="form" />
				:
				<input type="text" name="country" size="3" maxlength="2" class="form" />
				<BR />
				<BR />
				<input type="submit" value="Launch Task Force" class="form" />
							
													
				</table>
			</TD>
		</tr>
	</table>
</form>
<BR />
<BR />
<%
rsunits.Close 
strsql = "SELECT ETA, infantry, comandoes, jeeps, matilda1, matilda2, sherman, spitfire1, spitfire2, spitfire5, hurricane, lancaster, wellington, halifax, landingcraft, landingship, frigates, battleships, carrier, status, fromcountry, fromcontinent, tocountry, tocontinent, orderno FROM tblMovements WHERE username = '" & strusername & "' AND status > 0 AND status < 4 ORDER BY ETA ASC;"

rsunits.Open strsql, adocon

do while not rsunits.EOF 
%>
<center>
<table cellpadding="2" width="95%" cellspacing="0" border="0" style="border-top:solid; border-bottom:solid; border-right:solid; border-left:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline" align="center">
			<%
			select case rsunits("status")
				case 1
					Response.Write("Attacking " & rsunits("tocontinent") & ":" & rsunits("tocountry"))
				case 2
					Response.Write("Defending " & rsunits("tocontinent") & ":" & rsunits("tocountry"))
				case 3
					Response.Write("Returning Home")
			end select
			%>
			ETA: <% =rsunits("eta") %> Ticks 
		</TD>
	</TR>
	<TR>
		<TD class="norm" align="center">
		<table cellpadding="2" width="100%" border="0" cellspacing="0">
			<TR>
				<TD align="right" class="norm" width="30%">
					Infantry
				</TD>
				<TD class="norm" width="20%">
					<% =rsunits("infantry") %>
				</TD>
				<TD align="right" class="norm" width="30%">
					Commandos
				</TD>
				<TD class="norm" width="20%">
					<% =rsunits("comandoes") %>
				</TD>
			</TR>
			<TR>
				<TD align="right" class="norm">
					Jeeps
				</TD>
				<TD class="norm">
					<% =rsunits("jeeps") %>
				</TD>
				<TD align="right" class="norm">
					Matilda MkI
				</TD>
				<TD class="norm">
					<% =rsunits("matilda1") %>
				</TD>
			</TR>
			<TR>
				<TD align="right" class="norm">
					Matilda MkII
				</TD>
				<TD class="norm">
					<% =rsunits("matilda2") %>
				</TD>
				<TD align="right" class="norm">
					Sherman Tanks
				</TD>
				<TD class="norm">
					<% =rsunits("sherman") %>
				</TD>
			</TR>
			<TR>
				<TD align="right" class="norm">
					Supermarine Spitfire MkI
				</TD>
				<TD class="norm">
					<% =rsunits("spitfire1") %>
				</TD>
				<TD align="right" class="norm">
					Supermarine Spitfire MkII
				</TD>
				<TD class="norm">
					<% =rsunits("spitfire2") %>
				</TD>
			</TR>
			<TR>
				<TD align="right" class="norm">
					Supermarine Spitfire MkV
				</TD>
				<TD class="norm">
					<% =rsunits("spitfire5") %>
				</TD>
				<TD align="right" class="norm">
					Hawker Hurricane MkI
				</TD>
				<TD class="norm">
					<% =rsunits("hurricane") %>
				</TD>
			</TR>
			<TR>
				<TD align="right" class="norm">
					Halifax Bomber
				</TD>
				<TD class="norm">
					<% =rsunits("halifax") %>
				</TD>
				<TD align="right" class="norm">
					Wellington Bomber
				</TD>
				<TD class="norm">
					<% =rsunits("wellington") %>
				</TD>
			</TR>
			<TR>
				<TD align="right" class="norm">
					Lancaster Bomber
				</TD>
				<TD class="norm">
					<% =rsunits("lancaster") %>
				</TD>
				<TD align="right" class="norm">
					Landing Craft
				</TD>
				<TD class="norm">
					<% =rsunits("landingcraft") %>
				</TD>
			</TR>
			<TR>
				<TD align="right" class="norm">
					Landing Ship
				</TD>
				<TD class="norm">
					<% =rsunits("landingship") %>
				</TD>
				<TD align="right" class="norm">
					Frigates
				</TD>
				<TD class="norm">
					<% =rsunits("frigates") %>
				</TD>
			</TR>
			<TR>
				<TD align="right" class="norm">
					Battle Ships
				</TD>
				<TD class="norm">
					<% =rsunits("battleships") %>
				</TD>
				<TD align="right" class="norm">
					Aircraft Carriers
				</TD>
				<TD class="norm">
					<% =rsunits("carrier") %>
				</TD>
			</TR>
			</table>
			<BR />
			<BR />
			<%
			if rsunits("status") <> 3 then
			%>
			<A HREF="recal_force.asp?orderid=<% =rsunits("orderno") %>" onclick="return confirm('Are you sure you want to recall this task force?');">Recall</A>
			<%
			end if
			%>
		</TD>
	</TR>
</table>
<BR />
<BR />

<%
rsunits.MoveNext 
loop
rsunits.Close 

set rsunits = nothing
%>
</center>
<!---#include file="includes\footer.asp"--->
<!---#include file="includes\end_check.asp"--->