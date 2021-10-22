<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML><HEAD><TITLE>DSTAIR: A Decision Support Tool to Analyze Institutional Reform</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<META 
content="DSTAIR, Decision Support Tool, Institutional Reform, corruption" 
name=keywords>
<META 
content="DSTAIR: A Decision Support Tool to Analyze Institutional Reform" 
name=description>

</head>
<!--#include file ="conn.inc"-->
<!--#include file ="programs.asp"-->

<body>
<% if len(session("scenariosID")) = 0 then
	response.Redirect("default.asp")
end if %>
<%savedata%>
<%calculate(-1)%>
<% Response.Clear %>
<%response.redirect(Request.Form("ggoto"))%>

</body>
</html>