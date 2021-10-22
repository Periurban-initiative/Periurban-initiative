<LINK href="img/weblog2.css" type=text/css rel=stylesheet>
<% sub actools 
if len(session("scenariosID")) = 0 then
	response.Redirect("http://www.dstair.org/dstair4")
end if
if session("name")="default" then
	d = 1
else
	d = Request.QueryString("d")
end if	
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT count(*) as z FROM score " _
	& "INNER JOIN spheres ON score.spheresID = spheres.spheresID " _
	& "WHERE score.scenariosID = " & Session("scenariosID") & " AND (score.score <> -1);"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
if rs("z") > 0 then
	noq = false
else
	noq = true
end if
rs.close
set rs = nothing

Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT tooldata.toolID as toolID, tooldata.trigger as trigger, tools.title, tools.desc FROM tooldata " _
	& "INNER JOIN tools ON tools.toolID = tooldata.toolID WHERE tooldata.scenariosID = " & Session("scenariosID") _
	& " ORDER BY tooldata.trigger DESC,tools.toolID;"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
Set rs3 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM score " _
	& "INNER JOIN spheres ON score.spheresID = spheres.spheresID " _
	& "WHERE score.scenariosID = " & Session("scenariosID") & " ORDER BY spheres.ordr;"
rs3.Open strSQL, conn, adOpenForwardOnly, , adCmdText
%>
      <DIV class=floatcolmiddlehome> 
	  <img align="left" src="img/title_actools.gif"><a href="tool.asp?go=t&d=<%if d=1 then%>0<%else%>1<%end if%>"><img border=0 align="left" alt="Click to show or hide criteria used to trigger each tool" src="img/button_criteria.gif"></a><BR><BR>
	  <TABLE cellSpacing=0 cellPadding=4 border=2 bgcolor="#FF6666" width="100%"><TBODY><TR><TD>
<font color="#FFFFFF"><SPAN class=hed12px><strong>Currently selected analysis: <%=Session("scenariostitle")%> (<%=Session("scenarioscountry")%>)</strong></SPAN></FONT>
</TD></TR></TBODY></TABLE><BR>
<% if noq then %>
<font size=+1 color="#FF0000">NOTE: You have not yet answered any of the questions in the questionnaire.</font><BR>
<%end if%>
<form action="savetoolscriteria.asp" method="post" name="data">
<%
flip=0
while rs.EOF = false
Set rs2 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT spheres.spheresID, spheres.abbrev, toolcriteria.data FROM toolcriteria " _
	& "INNER JOIN spheres ON toolcriteria.spheresID = spheres.spheresID " _
	& "WHERE toolcriteria.toolID= " & rs("toolID") & " ORDER BY spheres.ordr;"
rs2.Open strSQL, conn, adOpenForwardOnly, , adCmdText
%>
<font size=-1><strong>
<%if (rs("trigger") = Session("numspheres")) and flip=0 and noq=false then
flip = 1%>
TRIGGERED ANTI-CORRUPTION TOOLS:<BR>
<%end if%>
<%if ((rs("trigger") < Session("numspheres")) or noq=true) and flip<2  then
flip=2%>
<BR>ALL ANTI-CORRUPTION TOOLS:<BR>
<%end if%>
</strong></font>
          <TABLE cellSpacing=0 cellPadding=5 width=579 bordercolor="#0E4B78" border=0><TBODY>
			<TR><TD>

          <TABLE cellSpacing=0 cellPadding=0 width=575 bgColor=#e47905 border=0 align=center><TBODY>
        <tr><td valign="top" bgcolor="#e47905" width="575" height="1"><font size="1"><img src="/wb/images/spacer.gif" height="1" width="520"></font></td></tr>
		  	<TR><TD>
			<TABLE cellSpacing=0 cellPadding=0 width=573 align=center  border=0><TBODY>
			<TR <%if (rs("trigger") = Session("numspheres")) then%>bgColor="#e47905"<%else%>bgColor="#BBBBBB"<%end if%>>
			<TD width=100 >&nbsp;&nbsp;<FONT size=-2 color="#FFFFFF">Tool #</FONT>&nbsp;&nbsp;<font size=+1 color="#FFFFFF"><strong><%=rs("toolID")%></strong></font></TD>
			<TD width=425 ><div align="left"><%if (rs("trigger") <> Session("numspheres")) then%><FONT size=-2 color="#FFFFFF">NOT TRIGGERED</FONT><%else%><FONT size=-2 color="#FFFFFF">TRIGGERED</FONT><%end if%></div></TD>
			<TD width=50 ><a href="toolhelp.asp?toolID=<%=rs("toolID")%>" target="_blank"><img border=0 src="img/info_icon.gif" alt="Click here for RESOURCES & CASE STUDIES pertaining to this tool"></a></TD>
			</TR></TBODY></TABLE>
			</TD></TR>
			</TBODY></TABLE>
			<TABLE cellSpacing=1 cellPadding=8 width=575 bgColor=#e47905 border=0 align=center><TBODY>
              <TR> 
				<TD valign="top" bgcolor="#EEEEEE"><SPAN class=hed12px><strong><%=rs("title")%></strong></SPAN>&nbsp;&#8212;&nbsp;<SPAN class=hed10px><%=rs("desc")%></SPAN></TD>
			  </TR>
			    
<%if d>0 then %>
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
					<FONT size=-2><strong>TOOL SELECTION CRITERIA</strong></FONT>
					<%if session("name")="default" then%><BR><font size=-2 color="#FF0000"><strong>EDIT MODE</strong></font><%end if%></td>
<%
rs2.movefirst
rs3.movefirst
for j = 1 to Session("numspheres")
		%>
		<td valign="top" align="center">
		<%if Session("name") <> "default" then%>
			<table width="100%" border=0 cellpadding="2" cellspacing="0" bgcolor="#FFFFFF">
			<tr><td align="center">
			<FONT size=-2><%=rs2("data")%></FONT>
			</td></tr></table>
		<%else%>
			<input type="text" name="tool_<%=rs("toolID")%>_<%=j%>"  size="3" maxlength="4" STYLE="font-size : 8pt" value="<%=rs2("data")%>">
		<%end if%>
		</td>
		<%
	rs2.movenext
	rs3.movenext
next 
%>
<td>&nbsp;</td></TR>
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
			<table width="100%" border="<%if rs3("score")>=rs2("data") then%>1<%else%>0<%end if%>" cellpadding="2" cellspacing="0" bgcolor="#FFFFFF" bordercolor="#990000">
			<tr><td align="center">
			<FONT size=-2><%if rs3("score")=-1 then%>NA<%else%><%=rs3("score")%><%end if%></FONT>
			</td></tr></table>
		</td>
			
	  <%rs2.movenext
	  rs3.movenext
	  next
		%>
<td>&nbsp;</td></TR>
					</TBODY></TABLE>
		</TD>
			  </TR>
<%end if%>			  
            </TBODY>
          </TABLE>
		</TD></TR></TBODY></TABLE>
<% rs.movenext
rs2.close
set rs2 = nothing
wend
rs.close
set rs = nothing
rs3.close
set rs3 = nothing

%>				
<BR>
<%if session("name")="default" then%><INPUT type="submit" value="SUBMIT CHANGES"><%end if%>
</FORM>
	  <BR><a href="tool.asp?go=r"><img align="right" src="img/button_next.gif" border=0 alt="Reports"></a>
      </DIV>

<% end sub %>