<!---SPELL CHECKED--->

<!---#include file="includes/ptop.asp"--->
<SCRIPT language="javascript">
	//<--hide script from older browsers
			
	//function to open a new window
			
	function openWin(url)
	{
		//open the new window
		window.open(url,'email', config='height=450,width=650,toolbar=no');
	}
						
	//--end hiding>
</script>
<table cellpadding="5" border="0" cellspacing="5" style="border-left:solid; border-right:solid; border-top:solid; border-bottom:solid; border-width:1; border-color:#FFFFFF;" width="95%">
	<TR>
		<TD class="tdheader" align="center">
			Register for Worlds of War II
		</TD>
	</TR>
	<TR>
		<TD class="norm">
		<%
		Dim rsSignup		'recordet to decide if signups are allowed
		Dim blnAllowed		'varible for if signup is allowed
		
		strSQL = "SELECT AllowedSignups FROM tblconfig;"
		
		'create the obects needed to open the dtaabase
		set adocon = server.CreateObject("ADODB.Connection")
		set rsSignup = server.CreateObject("ADODB.Recordset")
		
		adocon.ConnectionString = "Provider=SQLOLEDB;Server=neptune;User ID=WoWUsr;Password=password;Database=WoWII;"
		
		adocon.Open 
		
		rsSignup.Open strsql, adocon
		
		if rsSignup.EOF then
				blnAllowed = True
			else
				blnAllowed = cbool(rsSignup("allowedsignups"))
		end if
		
		rsSignup.Close 
		
		adocon.Close 
		
		set rssignup = nothing
		set adocon = nothing
		
		if blnallowed = true then
		%>
		
		To register for Worlds of War II you must have an active email address.
		<BR />
		<BR />
		If you have never activated your email using our activation system, click <A HREF="javascript:openWin('activate_email1.asp');">here</A>.
		<BR />
		<BR />
		Once you have activated your email, you have 72 hours to complete the registration process.
		<BR />
		<BR />
		<BR />
		<center>
			Once you have activated your email, click the button below to continue your signup.
			<BR />
			<BR />
			<form action="register2.asp" method="get">
				<input type="submit" value="Continue Registration" class="form">
			</form>
		</center>
		<%
		else
		%>
		Sorry but registration for Worlds of War II is currently closed, please try again later.
		<% 
		end if
		%>
		</TD>
	</TR>
</table>
<BR />
<BR />
	

<!---#include file="includes/pfooter.asp"--->