<%
'SPELL CHECKED

Dim objEmail
Dim rsEmail
Dim adocon
Dim strSQL

strsql = "SELECT tblemailqueue.* FROM tblemailqueue;"

set adocon = server.CreateObject("ADODB.Connection")
set rsEmail = server.CreateObject("ADODB.Recordset")

adocon.ConnectionString = "Provider=SQLOLEDB;Server=neptune;User ID=WoWUsr;Password=password;Database=WoWII;"

rsEmail.LockType = 3

rsEmail.CursorType = 2

adocon.Open 

rsEmail.Open strsql, adocon
if not rsEmail.EOF then
set objemail = server.CreateObject ("Jmail.SMTPMail")
with objEmail
	.ServerAddress ="127.0.0.1"
	.Sender = "no-reply@worldsofwar.co.uk"
	.AddRecipient rsemail("emailto")
	.Subject = rsemail("emailsubject")
	.HTMLBody = rsemail("emailbody")
			
	.Execute 
end with
rsEmail.Delete
end if
 
rsEmail.Close 

adocon.Close 

set rsemail = nothing
set adocon = nothing
%>
True