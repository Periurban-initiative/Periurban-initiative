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
strSQL = "SELECT questions.*, spheres.title " _
	& "FROM questions INNER JOIN spheres ON questions.spheresID = spheres.spheresID " _
	& "WHERE (((questions.questionsID) = " & request.querystring("qID") & "));"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
%>
<BR>
<img align="left" src="img/title_moreinfo.gif"><img src="img/info_icon.gif" width="38" height="38" align="left"><a href="javascript:window.close();"><img border=0 align="right" alt="Close this window" src="img/button_close.gif"></a><BR>
<BR>
<BR>
	  <TABLE cellSpacing=0 cellPadding=0 width=590 border=0 bgcolor="#339900"><TBODY>
	  <TR align=left><TD style="PADDING-RIGHT: 1px; PADDING-LEFT: 4px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px"><FONT size =-2 color="#FFFFFF"><img align="absmiddle" src="img/bullet_arrow_red2.gif">&nbsp;<STRONG><%=rs("title")%></STRONG></FONT></TD></TR>
	</TBODY></TABLE>	
	
        <TABLE cellSpacing=0 cellPadding=5 width=579 bordercolor="#0E4B78" border=0><TBODY>
		<TR><TD>

          <TABLE cellSpacing=0 cellPadding=0 width=575 bgColor=#e47905 align=center><TBODY>
        <tr><td valign="top" bgcolor="#e47905" width="575" height="1"><font size="1"><img src="/wb/images/spacer.gif" height="1" width="520"></font></td></tr>

		  	<TR><TD bgColor="#e47905"><FONT size=-2 color="#FFFFFF">&nbsp;&nbsp;Question</FONT>&nbsp;&nbsp;<font size=+1 color="#FFFFFF"><strong><%=rs("ordr")%></strong></font></TD>
			</TR></TBODY></TABLE>
			<TABLE cellSpacing=1 cellPadding=8 width=575 bgColor=#e47905 border=0 align=center><TBODY>
              <TR> 
				<TD valign="top" bgcolor="#EEEEEE"><font size="-2"><%=rs("text")%></FONT></TD>
			  </TR>
			  <TR>
			    <TD bgcolor="#FCE2C7">
				<FONT size=-2><strong>RATING </strong></FONT>
				<%if len(rs("scale_low") & rs("scale_high"))>0 then %>
					&nbsp;&nbsp;&nbsp;&nbsp;<FONT size=-2><strong><%=rs("scale_low")%>&nbsp;<img src="img/scale.gif">&nbsp;<%=rs("scale_high")%></strong></FONT>
				<%end if%>
				</TD>
			  </TR>			  
			  			  
            </TBODY>
          </TABLE><BR>
          <TABLE cellSpacing=0 cellPadding=0 width=575 bgColor="#3366FF" align=center><TBODY>
		  	<TR><TD width=400><strong><FONT size=-1 color="#FFFFFF">&nbsp;&nbsp;More information</FONT></strong></TD>
			<TD width=175>&nbsp;</TD>
			</TR></TBODY></TABLE>
        <TABLE cellSpacing=1 cellPadding=8 width=575 bgColor="#3366FF" border=0 align=center>
          <TBODY>
              <TR><TD bgcolor="#EEEEEE" valign="top"><font size=-1><%=rs("helpinfo")%></FONT></TD></TR>			  
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
