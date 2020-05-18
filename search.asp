<!---#include file="includes/ptop.asp"--->

<!---SPELL CHECKED--->

<table cellpadding="2" cellspacing="0" border="0" width="95%" style="border-top:solid; border-left:solid; border-right:solid; border-bottom:solid; border-color:#FFFFFF; border-width:1;">
	<TR>
		<TD class="tdheader" align="center">
			Search for News
		</TD>
	</TR>
	<TR>
		<TD class="norm">
			<center>
				<form action="search.asp" method="get">
					<table cellpadding="2" border="0" cellspacing="0" width="400">
						<TR>
							<TD class="norm">
								Search For
							</TD>
							<TD class="norm">
								<input type="text" size="20" name="term" maxlength="75" class="form" />
							</TD>
							<TD class="norm">
								<select name="type" size="1" class="form">
									<option value="1">News Article</option>
									<option value="2">Posted By</option>
								</select>
							</TD>
							<TD class="norm">
								<input type="submit" value="Search" class="form" />
							</TD>
						</TR>
					</table>
				</form>
			</center>
			<BR />
			<%
			if trim(Request.QueryString("term")) <> "" then
			%>
			<center>
				<Big>
					Search Results
				</big>
			</center>
			<%
			Dim rsAnnounce		'varible to store the recordset collection
			Dim strSearchTerm	'varible to hold the search term the user has entered
			
			'get the user input
			strSearchTerm = Request.QueryString("term")
			
			strSearchTerm = trim(strSearchTerm)
			
			'generate the SQL text
			strSQL = "SELECT tblAnnouncements.message, username, posted FROM tblannouncements WHERE "
			
			if Request.QueryString("type") = 1 then
					strSQL = strSQL & "message LIKE '%" & strsearchterm & "%' "
				else
					strsql = strsql & "username = '" & strsearchterm & "' "
			end if
			strsql = strsql & "ORDER by posted DESC;"
			
			'create the objects to open the database
			set adocon = server.CreateObject("ADODB.Connection")
			
			set rsannounce = server.CreateObject("ADODB.Recordset")
			
			adocon.ConnectionString = "Provider=SQLOLEDB;Server=neptune;User ID=WoWUsr;Password=password;Database=WoWII;"
			
			adocon.Open 
			
			rsAnnounce.Open strsql, adocon
			
			if rsAnnounce.EOF then
			%>
			Sorry there is no news that matches your search.
			<%
			else
			do while not rsAnnounce.EOF 
			%>
			<div align="left" class="norm">
				<% =rsannounce("message") %>
			</div>
			<BR />
			<div align="right" class="norm">
				<I>
					Posted by <% =rsannounce("username") %> on <% =rsannounce("posted") %>
				</I>
			</div>
			<BR />
			<BR />
			<BR />
			<%
			rsAnnounce.MoveNext 
			loop
			end if
			rsAnnounce.Close
			adocon.Close
			
			set rsannounce = nothing
			set adocon = nothing
			end if
			%>
		</TD>
	</TR>
</table>
							
<BR />
<!---#include file="includes/pfooter.asp"--->