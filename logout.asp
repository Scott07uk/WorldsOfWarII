<%
'SPELL CHECKED

Response.Cookies("WoWII")("username") = "0"
Response.Cookies("WoWII")("password") = "0"
Response.Cookies("WOWII").Expires = now()

Response.Redirect("default.asp")

%>