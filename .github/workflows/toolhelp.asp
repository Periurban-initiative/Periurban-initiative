<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>DSTAIR: A Decision Support Tool to Analyze Institutional Reform</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="description" content="DSTAIR: A Decision Support Tool to Analyze Institutional Reform">
<meta name="keywords" content="">
<LINK href="img/weblog2.css" type=text/css rel=stylesheet>
</head>
<!--#include file ="conn.inc"-->
<!--#include file ="getFileContents.asp"-->
<body>
<STYLE>
A:link {
	FONT-WEIGHT: bold; COLOR: #3300CC; TEXT-DECORATION: none
}
A:visited {
	FONT-WEIGHT: bold; COLOR: #3300CC; TEXT-DECORATION: none
}
A:active {
	FONT-WEIGHT: bold; COLOR: #3300CC; TEXT-DECORATION: none
}
A:hover {
	FONT-WEIGHT: bold; CURSOR: pointer; COLOR: #3300CC; TEXT-DECORATION: underline
}
</STYLE>
<% if len(session("scenariosID")) = 0 then
	response.Redirect("http://www.dstair.org/dstair4")
end if 
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT tools.* FROM tools " _
	& "WHERE tools.toolID = " & Request.QueryString("toolID") & ";"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
%>
<BR>
<img align="left" src="img/title_moreinfo.gif"><img src="img/info_icon.gif" width="38" height="38" align="left"><a href="javascript:window.close();"><img border=0 align="right" alt="Close this window" src="img/button_close.gif"></a><BR>
<BR>
<BR>
	
        <TABLE cellSpacing=0 cellPadding=5 width=579 bordercolor="#0E4B78" border=0><TBODY>
		<TR><TD>

          <TABLE cellSpacing=0 cellPadding=0 width=575 bgColor=#e47905 align=center><TBODY>
        <tr><td valign="top" bgcolor="#e47905" width="575" height="1"><font size="1"><img src="/wb/images/spacer.gif" height="1" width="520"></font></td></tr>

		  	<TR><TD bgColor="#e47905"><FONT size=-2 color="#FFFFFF">&nbsp;&nbsp;Tool #</FONT>&nbsp;&nbsp;<font size=+1 color="#FFFFFF"><strong><%=rs("toolID")%></strong></font></TD>
			</TR></TBODY></TABLE>
			<TABLE cellSpacing=1 cellPadding=8 width=575 bgColor=#e47905 border=0 align=center><TBODY>
              <TR> 
				<TD valign="top" bgcolor="#EEEEEE"><SPAN class=hed12px><strong><%=rs("title")%></strong></SPAN>&nbsp;&#8212;&nbsp;<SPAN class=hed10px><%=rs("desc")%></SPAN></TD>
			  </TR>
            </TBODY>
          </TABLE><BR>
          <TABLE cellSpacing=0 cellPadding=0 width=575 bgColor="#3366FF" align=center><TBODY>
		  	<TR><TD width=400><strong><FONT size=-1 color="#FFFFFF">&nbsp;&nbsp;Resources and Case Studies</FONT></strong></TD>
			<TD width=175>&nbsp;</TD>
			</TR></TBODY></TABLE>
        <TABLE cellSpacing=1 cellPadding=8 width=575 bgColor="#3366FF" border=0 align=center>
          <TBODY>
              <TR><TD bgcolor="#EEEEEE" valign="top"><%getfilecontents("casestudies/C" & request.QueryString("toolID"))%></TD></TR>			  
            </TBODY>
          </TABLE>
</td>
</tr>

            </TBODY>
          </TABLE>
<%
rs.close
set rs=nothing
%>
<BR>&nbsp;
</body>
</html>
