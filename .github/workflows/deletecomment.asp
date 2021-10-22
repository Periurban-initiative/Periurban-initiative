<!--#include file ="conn.inc"-->

<%
Response.Buffer = true
usersID = Session("usersID")
if len(usersID) = 0 then
	Response.Clear
	Session.Abandon
	response.redirect("default.asp")
end if
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM comments WHERE commentsID=" & Request.QueryString("commentsID") & ";"
rs.Open strSQL, conn, 1,3
if rs.EOF then
	Response.Clear
	response.redirect("default.asp?error=1")
end if
if (rs("usersID") <> session("usersID")) and (session("usersID") <> 0) then
	Response.Clear
	Session.Abandon
	response.redirect("default.asp")
end if	
rs.delete
rs.close
set rs = nothing

Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM questions WHERE questionsID=" & Request.querystring("qID") & ";"
rs.Open strSQL, conn, 1,3
rs("numcomments") = rs("numcomments") - 1
rs.update
rs.close
set rs = nothing
Response.Clear
response.redirect("comments.asp?qID=" & Request.querystring("qID"))
%>

