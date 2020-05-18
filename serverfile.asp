<%
'SPELL CHECKED

Dim adocon
Dim rsusers
Dim strCon
Dim rsCommon
Dim strSql

Set adoCon = Server.CreateObject("ADODB.Connection")

strCon = "Provider=SQLOLEDB;Server=neptune;User ID=WoWUsr;Password=password;Database=WoWII;"

adoCon.connectionstring = strCon

adocon.Open 

Set rsCommon = Server.CreateObject("ADODB.Recordset")

strSql = "SELECT tblaccounts.password FROM tblaccounts WHERE username = 'test';"

rsCommon.Open strSql, adocon

Response.Write(rsCommon("password"))

rsCommon.Close

adocon.Close

set rscommon = nothing
set adocon = nothing
%>