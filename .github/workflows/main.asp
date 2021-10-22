<LINK href="img/weblog2.css" type=text/css rel=stylesheet>
<%

sub main1
ii = Request.QueryString("i")
if len(Request.QueryString("i"))=0 then
	ii = 0
end if

%>
<script language="JavaScript">
function Submitb()
{
if(confirm ("ARE YOU SURE YOU WANT TO OVERWRITE YOUR SCENARIO WITH DEFAULT VALUES?"))
	location.replace("restoredefault.asp?sphere=<%=Request.QueryString("sphere")%>&t=<%=Request.QueryString("t")%>");
}
</script>
<%
sphere = Request.QueryString("sphere")
if len(Request.QueryString("scenariosID"))>0 then
	s=Request.QueryString("scenariosID")
else
	s = 0
end if
%>
<TABLE cellSpacing=0 cellPadding=6 width="100%" border=0><TBODY>
<TR><TD width="19%" align=left vAlign=top id=leftblocks name="leftblocks" bgcolor="#FFFFFF"> 
	<table width=100% cellspacing=0 cellpadding=4 border=1 bordercolor="#666699" bgcolor="#FFFFCC"><tbody>
	<tr><td align="center" bgcolor="#666699"><strong><font color="#FFFFFF" size=-2 face="Verdana, Arial, Helvetica, sans-serif">MY SCENARIO</font></strong></td></tr>
	<tr><td align="center"><font size=-2>ID = <%=Session("Name")%></font>&nbsp;</td></tr>
	<tr><td align="center"><font size=-2>Country = <%=Session("Title")%></font></td></tr>
	<tr><td align="center" bgcolor="#CCCCFF"><font size=-2><strong><a href="default.asp?editscenario=1"><font size=-2><strong>EDIT SCENARIO INFO</strong></font></a>
	<tr><td align="center" bgcolor="#CCCCFF"><font size=-2><strong>
	<% if Session("scenariosID") = 0 then %>
	YOU ARE CURRENTLY EDITING DEFAULT VALUES
	<% else %>
	<a href="javascript:Submitb();">RESTORE DEFAULTS</a>
	<%end if%>
	</strong></font></td></tr>
	<tr><td align="center" bgcolor="#CCCCFF"><a href="default.asp?scenarioslist=yes"><font size=-2><strong>VIEW OTHER USERS' SCENARIOS</strong></font></a></td></tr>
	<tr><td align="center" bgcolor="#CCCCFF"><a href="logincheck.asp?logoff=1"><font size=-2><strong>LOGOFF</strong></font></a></td></tr>
	
	</tbody></table><br>
	<table width=100% cellspacing=0 cellpadding=4 border=1 bordercolor="#666699" bgcolor="#FFFFCC"><tbody>
	<tr><td align="center" bgcolor="<%if s=0 then%>#666699<%else%>#CC0099<%end if%>">
	<strong><font color="#FFFFFF" size=-2 face="Verdana, Arial, Helvetica, sans-serif">TRUSTWORTHINESS MONITOR</font></strong></td></tr>
	<tr><td>
		<TABLE cellSpacing=1 cellPadding=4 width="100%" border=0><TBODY>
		<tr bgcolor="#CCCCFF"> 
		  <td width=75%><font color="#000000" size="-2"><strong>SPHERE</strong></font></td>
		  <td width=25%><font color="#000000" size="-2"><strong>SCORE</strong></font></td>
		</tr>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT spheres.spheresID, spheres.title, score.score, spheres.ordr as o " _
	 & "FROM score INNER JOIN spheres ON score.spheresID = spheres.spheresID " _
	 & "WHERE ((score.scenariosID=" 
if s=0 then
	strSQL = strSQL & Session("scenariosID")
else
	strSQL = strSQL & s
end if
strSQL = strSQL &  ")) ORDER BY spheres.ordr;"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
i = 0
while rs.EOF = false
i = i + 1
%>
                    <tr> 
                      <td width=75% <%if (rs("spheresID") = cint(Request.QueryString("sphere"))) then%>bgcolor="#000066"<%end if%>><font size="-2"><a href="default.asp?<%if s>0 then%>scenariosID=<%=s%>&<%end if%>sphere=<%=rs("spheresID")%>&t=<%=rs("title")%>&i=<%=ii%>"><%if (rs("spheresID") = cint(Request.QueryString("sphere"))) then%><font color="#99FF00"><%=rs("title")%></font><%else%><%=rs("title")%><%end if%></a></font></td>
                      <td width=25% <%if (rs("spheresID") = cint(Request.QueryString("sphere"))) then%>bgcolor="#000066"<%end if%>><font size="-2"><%if (rs("spheresID") = cint(Request.QueryString("sphere"))) then%><font color="#99FF00"><%if rs("score")=-1 then%>NA<%else%><%=rs("score")%><%end if%></font><%else%><%if rs("score")=-1 then%>NA<%else%><%=rs("score")%><%end if%><%end if%></font></td>
                    </tr>				
<%
rs.movenext
wend
rs.close
set rs = nothing
'get number of items triggered
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT count(*) AS c FROM tooldata WHERE scenariosID = " 
if s=0 then
	strSQL = strSQL & Session("scenariosID")
else
	strSQL = strSQL & s
end if
strSQL = strSQL & " AND trigger = " & Session("numspheres") & ";"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
trig = rs("c") & " of " & Session("numtools")
rs.close
set rs = nothing
%>
					
                </table>
		</td></tr>
		<tr><td align="center" bgcolor="#CCCCFF"><font color="#000000" size=-2 face="Verdana, Arial, Helvetica, sans-serif">click on a sphere<br>for details</font></td></tr>
		</table>
		<br>
	<table width=100% cellspacing=0 cellpadding=4 border=1 bordercolor="#666699" bgcolor="#FFFFCC"><tbody>
	<tr><td align="center" bgcolor="<%if s=0 then%>#666699<%else%>#CC0099<%end if%>"><strong>
      <font face="Verdana, Arial, Helvetica, sans-serif" size="-2" color="#FFFFFF">
      ANTI-CORRUPTION TOOLS</font></strong></td></tr>
	<tr><td align="center"><font size=-2><strong><%=trig%></strong> tools<br>
      exceed threshold trustworthiness scores</font></td></tr>
	<tr><td align="center" bgcolor="#CCCCFF"><font size=-2><strong><a href="default.asp?tools=1&d=0&t=1">REVIEW TOOLS</a></strong></font>	</td></tr>
	</tbody></table>
		
	</td>
	<TD class=sidebar1 id=middleblocks vAlign=top background=img/bg_grid.gif name="middleblocks">
		<%
		sphere = len(Request.QueryString("sphere"))
		if len(Request.QueryString("tools")) > 0 then
			if len(request.QueryString("n")) > 0 then
				bodycasestudies
			else
				bodytools
			end if
		else if len(request.QueryString("editscenario")) > 0 then
			body1
		else if len(request.QueryString("scenarioslist")) > 0 then
			bodylistscenarios
		else if sphere = 0 then
			body0
		else
			body
		end if
		end if
		end if
		end if
		%>
	</TD>
</TR>
</TBODY></TABLE>

<%
end sub
%>