<!--#include file ="conn.inc"-->
<%
Response.Buffer = true
scenariosID = Session("scenariosID")
if len(scenariosID) = 0 then
	Response.Clear
	Session.Abandon
	response.redirect("tool.asp")
end if
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM scenarios WHERE scenariosID=" & session("ScenariosID") & ";"
rs.Open strSQL, conn, 1,3
rs("title") = Request.Form("title")
rs("desc") = Request.Form("desc")
Session("Title") = rs("title")
Session("Desc") = rs("desc")
rs.update
rs.close
set rs = nothing
Response.Clear
session("new") = 0
response.redirect("tool.asp?go=s")
%>

