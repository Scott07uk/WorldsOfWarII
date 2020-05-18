<!---#include file="includes/check.asp"--->
<!---#include file="includes/top.asp"--->
<!---#include file="includes/stdFunctions.asp"--->

<!---SPELL CHECKED--->
<center>
<table width="95%" cellpadding="0" border="0" cellspacing="0" style="border-top:none; border-bottom:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD align="center" class="norm">
			<A HREF="javascript:openHelp('11');">Help With This Page</A>
		</TD>
	</TR>
</table>
</center>
<BR />
<%
Dim rsCon			'varible to hold collection for checking construction
Dim blnNewAllowed	'varible to decide if a new construction is allowed

strSQL = "SELECT curcon, conticks, ResCE, conSC, conDC, conWF, conHF, conLF, ResVGM, conRF, ResSF, ResMPV, conHMP, ConMF, ConSPF, ResBGP, ResMWP, ResMPC, ResRP, ResFPC, ResEP, ConBF, ResLB, ResFG, ConAA, ResSIT, ConMP, ResCS, ResE, ConSRC, conJF, conSF, conRS FROM tblAccounts WHERE username = '" & strUsername & "';"

'create the object to open the database

set rsCon = server.CreateObject("ADODB.Recordset")

rsCon.Open strsql, adocon
%>
<center>
<% select case Request.QueryString("error")
	case 1
		Response.Write("Sorry but you do not have enough money to start that construction. <BR /><BR />")
	case 2
		Response.Write("Sorry but you do not have enough iron to start that construction.  <BR /><BR />")
	case 3
		Response.Write("Sorry but you do not have enough wood to start that construction.  <BR /><BR />")
	case 4
		Response.Write("Sorry but you do not have enough food to start that construction.  <BR /><BR />")
	case 5
		Response.Write("The tick engine is currently processing, you can not start a new construction during this time, please try again in a few seconds.  <BR /><BR />")
end select %>
<table cellpadding="2" cellspacing="0" border="0" width="95%" style="border-top:solid; border-left:solid; border-right:solid; border-bottom:solid; border-width:1; border-color:#FFFFFF;">
	<TR>
		<TD class="tdheaderline" align="center">
			Current Construction
		</TD>
	</TR>
	<TR>
		<TD class="norm" align="center">
			<%
			if rscon("curcon") = "" then
			blnNewAllowed = True
			%>
			Your construction teams are standing by for your orders.  
			<%
			else
			blnNewAllowed = False
			%>
			Your construction teams are currently building
			<%
			Response.Write(rtnConName(rscon("curcon")))
			%>
			and will finish in <% =rscon("conticks") %> ticks.
			<%
			end if
			%>
		</TD>
	</TR>
</table>
<BR />
<BR />
<table align="center" class="norm" cellpadding="2" cellspacing="0" width="95%" style="border-top:solid; border-left:solid; border-right:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheaderline" width="">
			Construction Name
		</TD>
		<td class="tdheaderline" width="">
			Food
		</td>
		<td class="tdheaderline" width="">
			Wood
		</td>
		<td class="tdheaderline" width="">
			Iron
		</td>
		<td class="tdheaderline" width="">
			Money
		</td>
		<td class="tdheaderline" width="">
			Time
		</td>
		<td class="tdheaderline" width="">
			Status
		</td>
	</TR>
	<%
	if cint(rsCon("ResVGM")) = 3 then
	%>
	<TR height="27">
		<td class="all">
			Jeep Factory
		</td>
		<td class="all">
			800
		</TD>
		<TD class="all">
			250
		</TD>
		<TD class="all">
			1500
		</TD>
		<TD class="all">
			1250
		</td>
		<td class="all">
			12 ticks
		</td>
		<td class="all">
			<% if rscon("conJF") = 3 then %>
				Construction Complete
			<% elseif blnNewAllowed = true then %>
			<a href="con_engine.asp?start=1">Start</a>
			<% elseif rscon("CurCon") = "JF" then %>
			<% =rscon("conticks") %> Ticks Left
			<% else %>
			<font color="#FF0000">
				Other Construction in Progress
			</font>
			<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This builds the factory that produces jeeps.</i></small>
		</td>
	</tr>
	<%
	end if
	if cint(rsCon("ResSF")) = 3 then
	%>
	<TR height="27">
		<td class="all">
			Advanced Academy
		</td>
		<td class="all">
			500
		</TD>
		<TD class="all">
			500
		</TD>
		<TD class="all">
			750
		</TD>
		<TD class="all">
			1000
		</td>
		<td class="all">
			15 ticks
		</td>
		<td class="all">
			<% if rscon("conAA") = 3 then %>
				Construction Complete
			<% elseif blnNewAllowed = true then %>
			<a href="con_engine.asp?start=2">Start</a>
			<% elseif rsCon("CurCon") = "AA" then %>
			<% =rscon("conticks") %> Ticks Left
			<% else %>
			<font color="#FF0000">
				Other Construction in Progress
			</font>
			<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This building is used to train more special troops such as commandos.</i></small>
		</td>
	</tr>
	<%
	end if
	if (cint(rsCon("ResMPV")) = 3) AND (cint(rscon("ResVGM")) = 3) then
	 
	%>
	<TR height="27">
		<td class="all">
			Matilda Factory
		</td>
		<td class="all">
			850
		</TD>
		<TD class="all">
			500
		</TD>
		<TD class="all">
			1000
		</TD>
		<TD class="all">
			1500
		</td>
		<td class="all">
			18 ticks
		</td>
		<td class="all">
			<% if rscon("ConMF") = 3 then %>
				Construction Complete
			<% elseif blnNewAllowed = true then %>
			<a href="con_engine.asp?start=3">Start</a>
			<% elseif rscon("CurCon") = "MF" then %>
			<% =rscon("conticks") %> Ticks Left
			<% else %>
			<font color="#FF0000">
				Other Construction in Progress
			</font>
			<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This factory produces all marks of the Matilda tank.</i></small>
		</td>
	</tr>
	<%
	end if
	if (cint(rsCon("ResMPV")) = 3) AND (cint(rsCon("resVGM")) = 3) AND (cint(rsCon("ResBGP")) = 3) then
	%>
	<TR height="27">
		<td class="all">
			Sherman Tank Factory
		</td>
		<td class="all">
			1000
		</TD>
		<TD class="all">
			500
		</TD>
		<TD class="all">
			1500
		</TD>
		<TD class="all">
			2000
		</td>
		<td class="all">
			24 ticks
		</td>
		<td class="all">
			<% if rscon("conSF") = 3 then %>
				Construction Complete
			<% elseif blnNewAllowed = true then %>
			<a href="con_engine.asp?start=4">Start</a>
			<% elseif rsCon("curCon") = "SF" then %>
			<% =rscon("conticks") %> Ticks Left
			<% else %>
			<font color="#FF0000">
				Other Construction in Progress
			</font>
			<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This building is used to produce the Sherman tank.</i></small>
		</td>
	</tr>
	<%
	end if
	if (cint(rscon("ResMPC")) = 3) AND (cint(rscon("ResCE")) = 3) then
	%>
	<TR height="27">
		<td class="all">
			Spitfire Factory
		</td>
		<td class="all">
			1000
		</TD>
		<TD class="all">
			750
		</TD>
		<TD class="all">
			2500
		</TD>
		<TD class="all">
			3500
		</td>
		<td class="all">
			36 ticks
		</td>
		<td class="all">
			<% if rscon("conSPF") = 3 then %>
				Construction Complete
			<% elseif blnNewAllowed = true then %>
			<a href="con_engine.asp?start=5">Start</a>
			<% elseif rsCon("curCon") = "SPF" then %>
			<% =rscon("conticks") %> Ticks Left
			<% else %>
			<font color="#FF0000">
				Other Construction in Progress
			</font>
			<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This building is the factory used to manufacture all the marks of Spitfire.</i></small>
		</td>
	</tr>
	<%
	end if
	if cint(rscon("ResRP")) = 3 then
	%>
	<TR height="27">
		<td class="all">
			Rocket Factory
		</td>
		<td class="all">
			900
		</TD>
		<TD class="all">
			1250
		</TD>
		<TD class="all">
			1980
		</TD>
		<TD class="all">
			1000
		</td>
		<td class="all">
			20 ticks
		</td>
		<td class="all">
			<% if rscon("conRF") = 3 then %>
				Construction Complete
			<% elseif blnNewAllowed = true then %>
			<a href="con_engine.asp?start=6">Start</a>
			<% elseif rscon("curCon") = "RF" then %>
			<% =rscon("conticks") %> Ticks Left
			<% else %>
			<font color="#FF0000">
				Other Construction in Progress
			</font>
			<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This is the factory that manufactures rockets for use in the later version of Spitfire.</i></small>
		</td>
	</tr>
	<%
	end if
	if (cint(rscon("ResFPC")) = 3) AND (cint(rsCon("ResCE")) = 3) then 
	%>
	<TR height="27">
		<td class="all">
			Hawker Production Centre
		</td>
		<td class="all">
			800
		</TD>
		<TD class="all">
			800
		</TD>
		<TD class="all">
			1250
		</TD>
		<TD class="all">
			1700
		</td>
		<td class="all">
			28 ticks
		</td>
		<td class="all">
			<% if rscon("conHMP") = 3 then %>
				Construction Complete
			<% elseif blnNewAllowed = true then %>
			<a href="con_engine.asp?start=7">Start</a>
			<% elseif rsCon("CurCon") = "HMP" then %>
			<% =rscon("conticks") %> Ticks Left
			<% else %>
			<font color="#FF0000">
				Other Construction in Progress
			</font>
			<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This factory is used to produce the Hawker Hurricane.</i></small>
		</td>
	</tr>
	<%
	end if
	if (cint(rsCon("ResMPC")) = 3) AND (cint(rscon("ResCE")) = 3) then
	%>
	<TR height="27">
		<td class="all">
			Halifax Manufacturing Plant
		</td>
		<td class="all">
			1250
		</TD>
		<TD class="all">
			760
		</TD>
		<TD class="all">
			1300
		</TD>
		<TD class="all">
			2800
		</td>
		<td class="all">
			32 ticks
		</td>
		<td class="all">
			<% if rscon("conHF") = 3 then %>
				Construction Complete
			<% elseif blnNewAllowed = true then %>
			<a href="con_engine.asp?start=8">Start</a>
			<% elseif rscon("curcon") = "HF" then %>
			<% =rscon("conticks") %> Ticks Left
			<% else %>
			<font color="#FF0000">
				Other Construction in Progress
			</font>
			<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This factory is used to produce the Halifax Bomber.</i></small>
		</td>
	</tr>
	<%
	end if
	if (cint(rsCon("ResMPC")) = 3) AND (cint(rsCon("ResLB")) = 3) AND (cint(rsCon("ResCE")) = 3) then
	%>
	<TR height="27">
		<td class="all">
			Wellington Factory
		</td>
		<td class="all">
			1250
		</TD>
		<TD class="all">
			1200
		</TD>
		<TD class="all">
			2100
		</TD>
		<TD class="all">
			3500
		</td>
		<td class="all">
			32 ticks
		</td>
		<td class="all">
			<% if rscon("conWF") = 3 then %>
				Construction Complete
			<% elseif blnNewAllowed = true then %>
			<a href="con_engine.asp?start=9">Start</a>
			<% elseif rscon("curcon") = "WF" then %>
			<% =rscon("conticks") %> Ticks Left
			<% else %>
			<font color="#FF0000">
				Other Construction in Progress
			</font>
			<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This factory is used to produce the Wellington bomber.</i></small>
		</td>
	</tr>
	<%
	end if
	if (cint(rsCon("ResMPC")) = 3) AND (cint(rsCon("ResLB")) = 3) AND (cint(rsCon("ResCE")) = 3) then
	%>
	<TR height="27">
		<td class="all">
			Lancaster Factory
		</td>
		<td class="all">
			1500
		</TD>
		<TD class="all">
			1300
		</TD>
		<TD class="all">
			3300
		</TD>
		<TD class="all">
			4250
		</td>
		<td class="all">
			48 ticks
		</td>
		<td class="all">
			<% if rscon("conlf") = 3 then %>
				Construction Complete
			<% elseif blnNewAllowed = true then %>
			<a href="con_engine.asp?start=10">Start</a>
			<% elseif rscon("curcon") = "LF" then %>
			<% =rscon("conticks") %> Ticks Left
			<% else %>
			<font color="#FF0000">
				Other Construction in Progress
			</font>
			<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This factory is used to produce the long range Lancaster bomber.</i></small>
		</td>
	</tr>
	<%
	end if
	%>
	<TR height="27">
		<td class="all">
			Main Port
		</td>
		<td class="all">
			3500
		</TD>
		<TD class="all">
			2500
		</TD>
		<TD class="all">
			4500
		</TD>
		<TD class="all">
			7800
		</td>
		<td class="all">
			72 ticks
		</td>
		<td class="all">
			<% if rscon("conmp") = 3 then %>
				Construction Complete
			<% elseif blnNewAllowed = true then %>
			<a href="con_engine.asp?start=11">Start</a>
			<% elseif rscon("curcon") = "MP" then %>
			<% =rscon("conticks") %> Ticks Left
			<% else %>
			<font color="#FF0000">
				Other Construction in Progress
			</font>
			<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>This is the main port uses to produce all large ships.</i></small>
		</td>
	</tr>
	<TR height="27">
		<td class="all">
			Defence Control Centre
		</td>
		<td class="all">
			1750
		</TD>
		<TD class="all">
			750
		</TD>
		<TD class="all">
			1500
		</TD>
		<TD class="all">
			2500
		</td>
		<td class="all">
			24 ticks
		</td>
		<td class="all">
			<% if rscon("conDC") = 3 then %>
				Construction Complete
			<% elseif blnNewAllowed = true then %>
			<a href="con_engine.asp?start=12">Start</a>
			<% elseif rscon("curcon") = "DC" then %>
			<% =rscon("conticks") %> Ticks Left
			<% else %>
			<font color="#FF0000">
				Other Construction in Progress
			</font>
			<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>Co-ordinates all the activity of your fixed defence units such as AA guns and water mines.</i></small>
		</td>
	</tr>
	<%
	if cint(rscon("resFG")) = 3 then
	%>
	<TR height="27">
		<td class="all">
			Radar Station
		</td>
		<td class="all">
			250
		</TD>
		<TD class="all">
			500
		</TD>
		<TD class="all">
			1250
		</TD>
		<TD class="all">
			1500
		</td>
		<td class="all">
			15 ticks
		</td>
		<td class="all">
			<% if rscon("conRS") = 3 then %>
				Construction Complete
			<% elseif blnNewAllowed = true then %>
			<a href="con_engine.asp?start=13">Start</a>
			<% elseif rscon("curcon") = "RS" then %>
			<% =rscon("conticks") %> Ticks Left
			<% else %>
			<font color="#FF0000">
				Other Construction in Progress
			</font>
			<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>Provides information on enemy air movements, allowing AA guns to defend you.</i></small>
		</td>
	</tr>
	<%
	end if
	if (cint(rsCon("ConAA")) = 3) AND (cint(rsCon("ResSIT")) = 3) AND (cint(rsCon("ConMP")) = 3) AND (cint(rscon("ResRP")) = 3) AND (cint(rscon("ResCS")) = 3) AND (cint(rscon("ResE")) = 3) then
	%>
	<TR height="27">
		<td class="all">
			Spy Training Center
		</td>
		<td class="all">
			1500
		</TD>
		<TD class="all">
			600
		</TD>
		<TD class="all">
			2000
		</TD>
		<TD class="all">
			2500
		</td>
		<td class="all">
			24 ticks
		</td>
		<td class="all">
			<% if rscon("conSC") = 3 then %>
				Construction Complete
			<% elseif blnNewAllowed = true then %>
			<a href="con_engine.asp?start=15">Start</a>
			<% elseif rscon("curcon") = "SC" then %>
			<% =rscon("conticks") %> Ticks Left
			<% else %>
			<font color="#FF0000">
				Other Construction in Progress
			</font>
			<% end if %>
		</td>
	</TR>
	<tr>
		<td colspan="7" class="norm" style="border-bottom: solid; border-width: 1px;">
			<i><small>Allows the training of spies for SEO.</i></small>
		</td>
	</tr>
	<%
	end if
	%>
</table>
</center>
<%
rsCon.Close

set rsCon = nothing
%>		
<!---#include file="includes/footer.asp"--->			
<!---#include file="includes/end_check.asp"--->