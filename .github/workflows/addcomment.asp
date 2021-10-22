<!--#include file ="conn.inc"-->

<%
Response.Buffer = true
usersID = Session("usersID")
if len(usersID) = 0 then
	Response.Clear
	Session.Abandon
	response.redirect("default.asp")
end if
if len(Request.Form("x")) = 0 then
	x = " "
else
	x = Request.Form("x")
end if
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM comments;"
rs.Open strSQL, conn, 1,3
rs.addnew
rs("qID") = cint(Request.Form("qID"))
rs("comment") = x
rs("dt") = now()
rs("usersID") = session("usersID")
rs.update
rs.close
set rs = nothing
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM questions WHERE questionsID=" & Request.Form("qID") & ";"
rs.Open strSQL, conn, 1,3
rs("numcomments") = rs("numcomments") + 1
rs.update
rs.close
set rs = nothing
Response.Clear
response.redirect("comments.asp?qID=" & Request.Form("qID"))
%>

