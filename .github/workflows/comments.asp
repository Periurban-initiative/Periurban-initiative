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
<script language="JavaScript">
function Submita(x,y)
{
if(confirm ("Are you sure you want to delete this comment?"))
	location.replace("deletecomment.asp?commentsID=" + x + "&qID=" + y);
}
</script>
<% if len(session("usersID")) = 0 then
	response.Redirect("default.asp")
end if 
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT questions.*, spheres.title " _
	& "FROM questions INNER JOIN spheres ON questions.spheresID = spheres.spheresID " _
	& "WHERE (((questions.questionsID) = " & request.querystring("qID") & "));"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
%>
<BR>
<img align="left" src="img/title_usercomments.gif"><img src="img/memo_icon.gif" width="38" height="28" align="left"><a href="javascript:window.close();"><img border=0 align="right" alt="Close this window" src="img/button_close.gif"></a><BR>
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
            </TBODY>
          </TABLE><BR>
          <TABLE cellSpacing=0 cellPadding=0 width=575 bgColor="#FFCC00" align=center><TBODY>
		  	<TR><TD width=400><strong><FONT size=-1 color="#000000">&nbsp;&nbsp;<%=rs("numcomments")%>&nbsp;User comment<%if rs("numcomments")<>1 then%>s<%end if%></FONT></strong>&nbsp;&nbsp;<font size=+1 color="#FFFFFF">&nbsp;</font><TD>
			<TD width=175>&nbsp;</TD>
			</TR></TBODY></TABLE>
<%
rs.close
set rs=nothing
%>
<TABLE cellSpacing=1 cellPadding=8 width=575 bgColor="#FFCC00" border=0 align=center><TBODY>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT comments.qID, comments.*, users.name " _
	& "FROM comments INNER JOIN users ON comments.usersID = users.usersID " _
	& "WHERE (((comments.qID)=" & request.querystring("qID") & ")) ORDER BY comments.dt DESC;"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
i = 0
while rs.EOF = false
i = i + 1
%>
              <TR bgcolor="#EEEEEE">
			  	<TD width=5% align="center"><font size="+1"><strong><%=i%></strong></font>
					<%if Session("usersID")="0" or Session("usersID")=rs("usersID") then%>
					<br><a href="javascript:Submita(<%=rs("commentsID")%>,<%=request.querystring("qID")%>);"><img border=0 align="middle" src="img/trashcan.gif" title="Delete your comment"></a>
					<%end if%>	
				</TD>
	
			  	<TD width=65% valign="top"><font size="-2"><%=rs("comment")%></FONT></TD>
				<td width="30%" valign="top" align="center">
					<font size=-2><strong><%=rs("dt")%><br><%=rs("name")%></strong></font>
				</TD>
				  </TR>			  
<%
rs.movenext
wend
rs.close
set rs = nothing
%>
<TR bgcolor="#FFCC00" ><TD colspan="3"><img src="/wb/images/spacer.gif" height="1" width="1"></TD></TR>
<TR bgcolor="#FCE2C7" align="center">
<form action="addcomment.asp" method="post" name="data">
<input type="hidden" name="qID" value="<%=request.querystring("qID")%>">
<td width="5%" valign="top" align=center bgcolor="#FFCC00"><font size="+1"><strong>&nbsp;</strong></font></td>
<td width="65%" valign="top" align="left"><font size="-1"><strong>Add a new comment:
<TEXTAREA NAME="x" ROWS=8 COLS=40 STYLE="font-size : 10pt" value=""></textarea>
</strong></font><BR>	<input type=submit name=submit value=submit></td>
<td width="30%" bgcolor="#FFCC00" valign="top" align="center">
</td></form></tr>
<TR bgcolor="#FFCC00" ><TD colspan="3"><img src="/wb/images/spacer.gif" height="1" width="1"></TD></TR>
            </TBODY>
          </TABLE>
</td>
</tr>

            </TBODY>
          </TABLE>

</body>
</html>
