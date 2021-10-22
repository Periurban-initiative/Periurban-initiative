<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Response.Buffer = true %>
<html>
<head>
<title>DSTAIR: A Decision Support Tool to Analyze Institutional Reform</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="description" content="DSTAIR: A Decision Support Tool to Analyze Institutional Reform">
<meta name="keywords" content="">
<LINK href="style/style.css" type=text/css rel=stylesheet>

</head>
<!--#include file ="conn.inc"-->
<!--#include file ="programs.asp"-->

<body>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "select * from scenarios WHERE scenariosID=" & Request.Form("scenID") & ";"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
if rs.EOF = FALSE then
	session("scenariosID") = cint(Request.Form("scenID"))
	Session("scenariostitle") = rs("title")
	Session("scenarioscountry") = rs("country")
end if
rs.close
set rs = nothing
Response.Clear %>
<%response.redirect(Request.Form("changeanalysis"))%>

</body>
</html>