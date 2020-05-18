
<!---#include file="includes\check.asp"--->
<!---#include file="includes\top.asp"--->
<%
'SPELL CHECKED

Dim lngReportID

lngReportid = Request.QueryString("id")

if isnumeric(lngreportid) = false then
%>
You have not selected a valid report to view.
<%
else
Dim rsReport
set rsreport = server.CreateObject("ADODB.Recordset")

strsql = "SELECT * FROM tblattackresults WHERE username = '" & strusername & "' AND reportid = " & clng(lngreportid)

rsReport.Open strsql, adocon

if rsReport.EOF then
%>
The report you wanted to view does not exist or does not belong to you.
<%
else
%>
<BR />
<BR />
<center>
	<table cellpadding="2" border="0" cellspacing="0" width="95%" style="border-top:solid; border-right:solid; border-bottom:solid; border-left:solid; border-color:#FFFFFF; border-width:1;">
		<TR>
			<TD style="border-bottom:solid; border-color:#FFFFFF; border-width:1;">&nbsp;
				
			</TD>
			<TD colspan="2" class="tdheaderline" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				Attackers
			</TD>
			<TD colspan="2" class="tdheaderline" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				Defenders
			</TD>
			<TD colspan="2" class="tdheaderline" align="center">
				Your Forces
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Units
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				Sent
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				Lost
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				Sent
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				Lost
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				Sent
			</TD>
			<TD class="norm" align="center">
				Lost
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Infantry
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attinfantrys"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attinfantryl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("definfantrys"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("definfantryl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youinfantrys"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youinfantryl"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Commandos
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attcomandoess"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attcomandoesl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defcomandoess"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defcomandoesl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youcomandoess"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youcomandoesl"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Jeeps
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attjeepss"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attjeepsl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defjeepss"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defjeepsl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youjeepss"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youjeepsl"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Matilda MkI
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attmatilda1s"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attmatilda1"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defmatilda1s"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defmatilda1l"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youmatilda1s"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youmatilda1l"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Matilda MkII
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attmatilda2s"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attmatilda2"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defmatilda2s"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defmatilda2l"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youmatilda2s"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youmatilda2l"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Sherman Tank
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attshermans"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attshermanl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defshermans"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defshermanl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youshermans"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youshermanl"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Spitfire MkI
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attspitfire1s"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attspitfire1l"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defspitfire1s"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defspitfire1l"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youspitfire1s"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youspitfire1l"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Spitfire MkII
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attspitfire2s"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attspitfire2l"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defspitfire2s"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defspitfire2l"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youspitfire2s"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youspitfire2l"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Spitfire MkV
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attspitfire5s"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attspitfire5l"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defspitfire5s"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defspitfire5l"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youspitfire5s"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youspitfire5l"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Hawker Hurricane
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("atthurricanes"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("atthurricanel"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defhurricanes"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defhurricanel"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youhurricanes"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youhurricanel"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Halifax Bomber
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("atthalifaxs"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("atthalifaxl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				0
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				0
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youhalifaxs"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youhalifaxl"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Wellington Bomber
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attwellingtons"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attwellingtonl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				0
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				0
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youwellingtons"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youwellingtonl"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Lancaster Bomber
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attlancasters"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attlancasterl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				0
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				0
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youlancasters"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youlancasterl"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Landing Craft
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attlandingcrafts"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attlandingcraftl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				0
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				0
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youlandingcrafts"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youlandingcraftl"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Landing Ships
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attlandingships"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attlandingshipl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				0
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				0
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youlandingships"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youlandingshipl"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Frigates
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attfrigates"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attfrigatesl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("deffrigatess"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("deffrigatesl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youfrigatess"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youfrigatesl"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Battleships
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attbattleships"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attbattleshipl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defbattleships"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defbattleshipl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youbattleships"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youbattleshipl"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Carriers
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attcarriers"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("attcarrierl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defcarriers"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defcarrierl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youcarriers"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youcarrierl"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Pillboxes
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				0
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				0
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defpilboxess"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defpilboxesl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youpilboxes"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youpilboxesl"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				AA Guns
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				0
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				0
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defaagunss"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defaagunl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youaaguns"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youaagunl"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Barrage Balloons
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				0
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				0
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defballoonss"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defballoonsl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youballons"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youballoonl"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Water Mines
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				0
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				0
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defminess"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defminesl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youmines"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youminel"),0) %>
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Turrets
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-left:solid; border-width:1; border-color:#FFFFFF;">
				0
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				0
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defturretss"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("defturretsl"),0) %>
			</TD>
			<TD class="norm" align="center" style="border-right:solid; border-width:1; border-color:#FFFFFF;">
				<% =formatnumber(rsreport("youturrets"),0) %>
			</TD>
			<TD class="norm" align="center">
				<% =formatnumber(rsreport("youturretl"),0) %>
			</TD>
		</TR>
	</table>
	<BR />
	<BR />
	<table cellpadding="2" border="0" width="95%" cellspacing="0" style="border-top:solid; border-bottom:solid; border-right:solid; border-left:solid; border-color:#FFFFFF; border-width:1;">
		<TR>
			<TD class="tdheader" colspan="5" align="center">
				Other Losses
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				Farms
			</TD>
			<TD class="norm">
				Mines
			</TD>
			<TD class="norm">
				Sawmills
			</TD>
			<TD class="norm">
				Refineries
			</TD>
			<TD class="norm">
				Land
			</TD>
		</TR>
		<TR>
			<TD class="norm">
				<% =formatnumber(rsreport("farmslost"),0) %>
			</TD>
			<TD class="norm">
				<% =formatnumber(rsreport("mineslost"),0) %>
			</TD>
			<TD class="norm">
				<% =formatnumber(rsreport("millslost"),0) %>
			</TD>
			<TD class="norm">
				<% =formatnumber(rsreport("refinerieslost"),0) %>
			</TD>
			<TD class="norm">
				<% =formatnumber(rsreport("landlost"),0) %>
			</TD>
		</TR>
	</table>

<%
end if
rsReport.Close 

set rsreport = nothing
end if
%>	
<!---#include file="includes\footer.asp"--->
<!---#include file="includes\end_check.asp"--->
