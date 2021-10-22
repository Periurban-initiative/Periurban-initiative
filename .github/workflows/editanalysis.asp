<!--#include file ="conn.inc"-->
<%
Response.Buffer = true
scenariosID = Request.Form("scenario")
if len(scenariosID) = 0 then
	Response.Clear
	Session.Abandon
	response.redirect("tool.asp?go=s")
end if
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM scenarios WHERE scenariosID=" & scenariosID & ";"
rs.Open strSQL, conn, 1,3
rs("title") = Request.Form("title")
rs("desc") = Request.Form("desc")
rs("country") = Request.Form("country")
rs.update
rs.close
set rs = nothing
Response.Clear
session("new") = 0
response.redirect("tool.asp?go=s")
%>

