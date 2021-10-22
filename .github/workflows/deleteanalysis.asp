<!--#include file ="conn.inc"-->
<%
Response.Buffer = true
if (Session("usersID") = -1) or (len(Session("usersID")) = 0) then
	Response.Clear
	Session.Abandon
	response.redirect("default.asp")
end if
scenID = Request.QueryString("scenID")
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM scenarios WHERE scenariosID=" & scenID & ";"
rs.Open strSQL, conn, 1,3
if rs.EOF=FALSE then
	rs.delete
end if
set rs = nothing
' delete all data 
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "select * from data WHERE scenariosID=" & scenID & ";"
rs.Open strSQL, conn, 1,3
while rs.EOF = false
	rs.delete
	rs.movenext
wend	
rs.close
set rs = nothing

' delete all scores
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM score WHERE scenariosID=" & scenID & ";"
rs.Open strSQL, conn, 1,3
while rs.EOF = false
	rs.delete
	rs.movenext
wend
rs.close
set rs = nothing

' delete all tooldata
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM tooldata WHERE scenariosID=" & scenID & ";"
rs.Open strSQL, conn, 1,3
while rs.EOF = false
	rs.delete
	rs.movenext
wend
rs.close
set rs = nothing


Response.Clear
session("new") = 0
response.redirect("tool.asp?go=s")
%>

