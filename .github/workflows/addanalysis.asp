<!--#include file ="conn.inc"-->
<%
Response.Buffer = true

Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM scenarios;"
rs.Open strSQL, conn, 1,3
rs.addnew
rs("usersID")=Session("usersID")
rs("title") = Request.Form("title")
rs("desc") = Request.Form("desc")
rs("country") = Request.Form("country")
rs.update
scenID = rs("scenariosID")
rs.close
set rs = nothing

Set rs2 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM data;"
rs2.Open strSQL, conn, 1,3
Set rs3 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM questions;"
rs3.Open strSQL, conn, adOpenForwardOnly, , adCmdText
while rs3.EOF = false
	rs2.addnew
	rs2("scenariosID") = scenID
	rs2("questionsID") = rs3("questionsID")
	rs2("rating") = 0
	rs2.update
	rs3.movenext
wend
rs2.close
set rs2 = nothing
rs3.close
set rs3 = nothing

'add empty score records for new scenario
Set rs2 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM score;"
rs2.Open strSQL, conn, 1,3
for i = 1 to Session("numspheres")
	rs2.addnew
	rs2("scenariosID") = scenID
	rs2("spheresID") = i
	rs2("score") = -1
	rs2.update
next
rs2.close
set rs2 = nothing

'add empty tooldata records for new scenario
Set rs2 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM tooldata;"
rs2.Open strSQL, conn, 1,3
for i = 1 to Session("numtools")
	rs2.addnew
	rs2("scenariosID") = scenID
	rs2("toolID") = i
	rs2("trigger") = Session("numspheres")
	rs2.update
next
rs2.close
set rs2 = nothing

Response.Clear
session("new") = 0
response.redirect("tool.asp?go=s")
%>

