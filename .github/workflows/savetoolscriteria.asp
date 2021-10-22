<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Response.Buffer = true %>
<html>
<head>
<title>DSTAIR: A Decision Support Tool to Analyze Institutional Reform</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="description" content="DSTAIR: A Decision Support Tool to Analyze Institutional Reform">
<meta name="keywords" content="">

</head>
<!--#include file ="conn.inc"-->
<!--#include file ="programs.asp"-->

<body>
<%savetoolscriteria%>
<% Response.Clear %>
<%response.redirect("tool.asp?go=t")%>

</body>
</html>