<LINK href="img/weblog2.css" type=text/css rel=stylesheet>
<% sub results
if len(session("usersID")) = 0 then
	response.Redirect("default.asp")
end if
if session("name")="default" then
	d = 1
else
	d = Request.QueryString("d")
end if	
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT tooldata.toolID as toolID, tooldata.trigger as trigger, tools.title, tools.desc FROM tooldata " _
	& "INNER JOIN tools ON tools.toolID = tooldata.toolID WHERE tooldata.scenariosID = " & Session("scenariosID") _
	& " ORDER BY tools.toolID;"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
Set rs2 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT spheres.spheresID, spheres.abbrev, toolcriteria.data FROM toolcriteria " _
	& "INNER JOIN spheres ON toolcriteria.spheresID = spheres.spheresID " _
	& "WHERE toolcriteria.toolID= " & rs("toolID") & " ORDER BY spheres.ordr;"
rs2.Open strSQL, conn, adOpenForwardOnly, , adCmdText
Set rs3 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM score " _
	& "INNER JOIN spheres ON score.spheresID = spheres.spheresID " _
	& "WHERE score.scenariosID = " & Session("scenariosID") & " ORDER BY spheres.ordr;"
rs3.Open strSQL, conn, adOpenForwardOnly, , adCmdText
%>
      <DIV class=floatcolmiddlehome> 
	  <img align="left" src="img/title_results.gif"><BR><BR>
	  <TABLE cellSpacing=0 cellPadding=4 border=2 bgcolor="#FF6666" width="100%"><TBODY><TR><TD>
<font color="#FFFFFF"><SPAN class=hed12px><strong>Currently selected analysis: <%=Session("scenariostitle")%> (<%=Session("scenarioscountry")%>)</strong></SPAN></FONT>
</TD></TR></TBODY></TABLE><BR>	  
          <TABLE cellSpacing=0 cellPadding=5 width=579 bordercolor="#0E4B78" border=0><TBODY>
			<TR><TD>

          <TABLE cellSpacing=0 cellPadding=0 width=575 bgColor=#e47905 border=0 align=center><TBODY>
        <tr><td valign="top" bgcolor="#e47905" width="575" height="1"><font size="1"><img src="/wb/images/spacer.gif" height="1" width="520"></font></td></tr>
		  	<TR><TD>
			<TABLE cellSpacing=0 cellPadding=0 width=573 align=center  border=0><TBODY>
			<TR bgColor="#e47905">
			<TD>&nbsp;&nbsp;<FONT size=-1 color="#FFFFFF"><STRONG>Summary of Legitimacy Scores</STRONG></FONT></TD>
			</TR></TBODY></TABLE>
			</TD></TR>
			</TBODY></TABLE>
			<TABLE cellSpacing=1 cellPadding=8 width=575 bgColor=#e47905 border=0 align=center><TBODY>
			  <TR>
				<TD bgcolor="#FCE2C7">
					<TABLE><TBODY>
					<TR>
					<TD width=300 align="center">&nbsp;</TD>
<% for j = 1 to Session("numspheres") %>
					<td width="40" align="center"><SPAN class=hed10px><%=ucase(rs2("abbrev"))%></SPAN></td>
<% 
rs2.movenext
next %>	
<td width=50>&nbsp;</td>				
					</TR>
					<TR>
					<td valign="top">
					<FONT size=-2><strong>LEGITIMACY SCORE</strong></FONT>
					</td>
<%
		rs2.movefirst
		rs3.movefirst
		%>
		
<%	  for i = 1 to 9%>
	  
		<td valign="top" align="center">
			<table width="100%" border=0 cellpadding="2" cellspacing="0" bgcolor="#FFFFFF">
			<tr><td align="center">
			<FONT size=-2><%if rs3("score")=-1 then%>NA<%else%><%=rs3("score")%><%end if%></FONT>
			</td></tr></table>
		</td>
			
	  <%rs2.movenext
	  rs3.movenext
	  next
		%>
<td>&nbsp;</td></TR>
<TR><TD colspan=10><BR><font size=-2><strong>NOTES:</strong><BR>HIGH SCORES indicate a high level of legitimacy.  Reforms that work through this sphere will have a high likelihood of success.<BR><BR>
LOW SCORES indicate an illegitimate and probably corrupt sphere of the government.  It is unlikely that any reforms which work through this sphere will be of much use.
</font></TD></TR>
					</TBODY></TABLE>
		</TD>
			  </TR>
            </TBODY>
          </TABLE>
		</TD></TR></TBODY></TABLE>
<% rs.movenext
rs2.close
set rs2 = nothing
rs3.close
set rs3 = nothing

%>				
<BR>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT tooldata.toolID as toolID, tooldata.trigger as trigger, tools.title, tools.desc FROM tooldata " _
	& "INNER JOIN tools ON tools.toolID = tooldata.toolID WHERE tooldata.scenariosID = " & Session("scenariosID") _
	& " ORDER BY tools.toolID;"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
Set rs2 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT spheres.spheresID, spheres.abbrev, toolcriteria.data FROM toolcriteria " _
	& "INNER JOIN spheres ON toolcriteria.spheresID = spheres.spheresID " _
	& "WHERE toolcriteria.toolID= " & rs("toolID") & " ORDER BY spheres.ordr;"
rs2.Open strSQL, conn, adOpenForwardOnly, , adCmdText

Set rs4 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM scenarios WHERE usersID = " & Session("usersID") & " ORDER BY scenariosID;"
rs4.Open strSQL, conn, adOpenForwardOnly, , adCmdText
%>
          <TABLE cellSpacing=0 cellPadding=5 width=579 bordercolor="#0E4B78" border=0><TBODY>
			<TR><TD>

          <TABLE cellSpacing=0 cellPadding=0 width=575 bgColor=#e47905 border=0 align=center><TBODY>
        <tr><td valign="top" bgcolor="#e47905" width="575" height="1"><font size="1"><img src="/wb/images/spacer.gif" height="1" width="520"></font></td></tr>
		  	<TR><TD>
			<TABLE cellSpacing=0 cellPadding=0 width=573 align=center  border=0><TBODY>
			<TR bgColor="#e47905">
			<TD>&nbsp;&nbsp;<FONT size=-1 color="#FFFFFF"><STRONG>Comparison of Legitimacy Scores With All Analyses</STRONG></FONT></TD>
			</TR></TBODY></TABLE>
			</TD></TR>
			</TBODY></TABLE>
			<TABLE cellSpacing=1 cellPadding=8 width=575 bgColor=#e47905 border=0 align=center><TBODY>
			  <TR>
				<TD bgcolor="#FCE2C7">
					<TABLE><TBODY>
					<TR>
					<TD width=300 align="center">&nbsp;</TD>
<% for j = 1 to Session("numspheres") %>
					<td width="40" align="center"><SPAN class=hed10px><%=ucase(rs2("abbrev"))%></SPAN></td>
<% 
rs2.movenext
next %>	
<td width=50>&nbsp;</td>				
					</TR>
<%
While rs4.EOF = FALSE
Set rs3 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM score " _
	& "INNER JOIN spheres ON score.spheresID = spheres.spheresID " _
	& "WHERE score.scenariosID = " & rs4("scenariosID") & " ORDER BY spheres.ordr;"
rs3.Open strSQL, conn, adOpenForwardOnly, , adCmdText
%>		
					<TR>
					<td valign="top">
					<FONT size=-2><strong><%=left((rs4("title")),19)%><BR>(<%if len(rs4("country"))>0 then %><%=left((rs4("country")),19)%><%else%>N/A<%end if%>)</strong></FONT><BR><BR>
					</td>
<%
		rs2.movefirst
		rs3.movefirst
		%>
		
<%	  for i = 1 to 9%>
	  
		<td valign="top" align="center">
			<table width="100%" border=0 cellpadding="2" cellspacing="0" bgcolor="#FFFFFF">
			<tr><td align="center">
			<FONT size=-2><%if rs3("score")=-1 then%>NA<%else%><%=rs3("score")%><%end if%></FONT>
			</td></tr></table>
		</td>
			
	  <%rs2.movenext
	  rs3.movenext
	  next
		%>
<td>&nbsp;</td></TR>
<%
rs4.movenext
wend
%>
					</TBODY></TABLE>
		</TD>
			  </TR>
            </TBODY>
          </TABLE>
		</TD></TR></TBODY></TABLE>
<% rs.movenext
rs2.close
set rs2 = nothing
rs3.close
set rs3 = nothing

%>				

	  <BR><a href="tool.asp?go=t"><img align="right" src="img/button_next.gif" border=0 alt="STEP 4: Tools"></a>

      </DIV>

<% end sub %>