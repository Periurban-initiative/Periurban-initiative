
<% sub menu %>
<SCRIPT language=JavaScript>
<!--
 
	if (document.images) {
<% For x = 1 to 9 %>
		var tab<%=x%>On = new Image();
		var tab<%=x%>Off = new Image();
		tab<%=x%>On.src = "img/tabover_<%=x%>.gif";
		tab<%=x%>Off.src = "img/tab_<%=x%>.gif";
<% next %>
<% For Each x in array("welcome", "instructions", "country", "questionaire", "results", "tools", "reports", "logoff", "xhome", "about", "xdemo", "resources", "contact") %>
		var <%=x%>On = new Image();
		var <%=x%>Off = new Image();
		<%=x%>On.src = "img/menuover_<%=x%>.gif";
		<%=x%>Off.src = "img/menu_<%=x%>.gif";
<% next %>
	}
	function act(imgName) {
		if (document.images)
		  {document[imgName].src = eval(imgName + "On.src");
  		  }
	}

	function inact(imgName) {
		if (document.images)
		  {document[imgName].src = eval(imgName + "Off.src");
   		  }
	}

//-->
</SCRIPT>
<%go = Request.QueryString("go")
if len(go)=0 then
	go = "w"
end if
%>
<%sphere = cint(Request.QueryString("sphere"))%>
  <DIV id=leftcolumn>
    <TABLE width=150 cellspacing="0" border="0" cellpadding="0">
      <TBODY><TR>
	  <td valign="top">
	  <img src="img/menu2.gif"><BR>
  <%if go="w" then%><img border="0" src="img/menuon_welcome.gif"><%else%><a href="javascript:Submita('tool.asp?go=w')" onMouseOver="act('welcome')" onMouseOut="inact('welcome')"><img border="0" src="img/menu_welcome.gif" name="welcome" id="welcome"></a><%end if%><BR>
 <%if go="i" then%><img border="0" src="img/menuon_instructions.gif"><%else%><a href="javascript:Submita('tool.asp?go=i')" onMouseOver="act('instructions')" onMouseOut="inact('instructions')"><img border="0" src="img/menu_instructions.gif" name="instructions" id="instructions"></a><%end if%><BR>
 <%if go="s" then%><img border="0" src="img/menuon_country.gif"><%else%><a href="javascript:Submita('tool.asp?go=s')" onMouseOver="act('country')" onMouseOut="inact('country')"><img border="0" src="img/menu_country.gif" name="country" id="country"></a><%end if%><BR>
  <%if go="q" then%><img border="0" src="img/menuon_questionaire.gif"><%else%><a href="javascript:Submita('tool.asp?go=q&sphere=1&t=CONSTITUTION')" onMouseOver="act('questionaire')" onMouseOut="inact('questionaire')"><img border="0" src="img/menu_questionaire.gif"  name="questionaire" id="questionaire"></a><%end if%><BR>
<%if go="q" then%>
    <TABLE width=152 cellspacing="0" border="0" cellpadding="0"><TBODY>
       <TR bgColor="<%if go="q" then%>#339900<%else%>#6699cc<%end if%>"> 
          <TD style="PADDING-RIGHT: 1px; PADDING-LEFT: 15px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px"> 
<form action="calc.asp" method="post" name="data">
<font color="#e9f4ff" size=-2 >
Switch to analysis: <a href="javascript:Submita('tool.asp?go=q&sphere=1&t=CONSTITUTION')"><img src="img/go_button.gif" border="0"></a><BR>
<select name="analysis" STYLE="font-size : 10.5px; font-family :'Verdana'">
<OPTION selected VALUE="<%=Session("scenariosID")%>"><%=left(Session("scenariostitle"),18)%>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "select * from scenarios WHERE usersID=" & Session("usersID") & " ORDER BY scenariosID;"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
while rs.EOF = false
if rs("scenariosID") <> cint(Session("scenariosID")) then
%>
<OPTION VALUE="<%=rs("scenariosID")%>"><%=left(rs("title"),18)%>
<%
end if
rs.movenext
wend
rs.close
set rs = nothing
%>
</SELECT>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM spheres ORDER BY spheres.ordr;"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
Set rs2 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT questions.spheresID, Sum(Abs([rating]>0)) AS t, Count(questions.questionsID) AS c " _
	& "FROM questions INNER JOIN data ON questions.questionsID = data.questionsID " _
	& "WHERE (((data.scenariosID)=" & Session("scenariosID") & ")) " _
	& "GROUP BY questions.spheresID;"
rs2.Open strSQL, conn, adOpenForwardOnly, , adCmdText
%>
<TABLE cellpadding="0" cellspacing="0"><TBODY>
<TR>
<TD width=38 align="center"><U>Incl</U></TD>
<TD><U>Sphere</U></TD>
<%
for i = 1 to 9
%>
<TR>
<TD width=38 align="center"><%=rs2("t")%>/<%=rs2("c")%></TD>
<TD><%if go="q" then%><a href="javascript:Submita('tool.asp?go=q&sphere=<%=rs("spheresID")%>&t=<%=rs("title")%>')"><%=rs("title")%></a><%else%><%=rs("title")%><%end if%></TD>
</TR>
<%
rs.movenext
rs2.movenext
next
rs.close
set rs = nothing
rs2.close
set rs2 = nothing
 %>
</TBODY></TABLE>
</TD></TR></TBODY></TABLE>
</font>
<%end if%>
<%'get number of items triggered
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT count(*) AS c FROM tooldata WHERE scenariosID = " & Session("scenariosID") & " AND trigger = " & Session("numspheres") & ";"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
trig = rs("c") & " of " & Session("numtools")
rs.close
set rs = nothing
%>
 <%if go="e" then%><img border="0" src="img/menuon_results.gif"><%else%><a href="javascript:Submita('tool.asp?go=e')" onMouseOver="act('results')" onMouseOut="inact('results')"><img border="0" src="img/menu_results.gif"  name="results" id="results"></a><%end if%><BR>
<%if go="e" then%>
    <TABLE width=152 cellspacing="0" border="0" cellpadding="0"><TBODY>
       <TR bgColor="<%if go="e" then%>#339900<%else%>#6699cc<%end if%>"> 
          <TD style="PADDING-RIGHT: 1px; PADDING-LEFT: 15px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px"> 

<form action="changeanalysis.asp" method="post" name="changeanalysis">
<font color="#e9f4ff" size=-2>
Switch to analysis: <a href="javascript:Submitaa()"><img src="img/go_button.gif" border="0"></a></FONT><BR>
<input type="hidden" name="changeanalysis" value="tool.asp?go=e">
<select name="scenID" STYLE="font-size : 10.5px; font-family :'Verdana'">
<OPTION selected VALUE="<%=Session("scenariosID")%>"><%=left(Session("scenariostitle"),18)%>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "select * from scenarios WHERE usersID=" & Session("usersID") & " ORDER BY scenariosID;"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
while rs.EOF = false
if rs("scenariosID") <> cint(Session("scenariosID")) then
%>
<OPTION VALUE="<%=rs("scenariosID")%>"><%=left(rs("title"),18)%>
<%
end if
rs.movenext
wend
rs.close
set rs = nothing
%>
</SELECT></FORM>
</TD></TR></TBODY></TABLE>
<%end if%>

 <%if go="t" then%><img border="0" src="img/menuon_tools.gif"><%else%><a href="javascript:Submita('tool.asp?go=t')" onMouseOver="act('tools')" onMouseOut="inact('tools')"><img border="0" src="img/menu_tools.gif"  name="tools" id="tools"><%end if%></a><BR>
<%if go="t" then%>
    <TABLE width=152 cellspacing="0" border="0" cellpadding="0"><TBODY>
       <TR bgColor="<%if go="t" then%>#339900<%else%>#6699cc<%end if%>"> 
          <TD style="PADDING-RIGHT: 1px; PADDING-LEFT: 15px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px">
<form action="changeanalysis.asp" method="post" name="changeanalysis">
<font color="#e9f4ff" size=-2>
Switch to analysis: <a href="javascript:Submitaa()"><img src="img/go_button.gif" border="0"></a></FONT><BR>
<input type="hidden" name="changeanalysis" value="tool.asp?go=t">
<select name="scenID" STYLE="font-size : 10.5px; font-family :'Verdana'">
<OPTION selected VALUE="<%=Session("scenariosID")%>"><%=left(Session("scenariostitle"),18)%>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "select * from scenarios WHERE usersID=" & Session("usersID") & " ORDER BY scenariosID;"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
while rs.EOF = false
if rs("scenariosID") <> cint(Session("scenariosID")) then
%>
<OPTION VALUE="<%=rs("scenariosID")%>"><%=left(rs("title"),18)%>
<%
end if
rs.movenext
wend
rs.close
set rs = nothing
%>
</SELECT></FORM>
<font color="#e9f4ff" size=-2 ><%=trig%> TRIGGERED</font>
</TD></TR></TBODY></TABLE>
<%end if%>
<%if go="r" then%><img border=0 src="img/menuon_reports.gif"><%else%><a href="tool.asp?go=r" onMouseOver="act('reports')" onMouseOut="inact('reports')"><img border=0 src="img/menu_reports.gif" name="reports" id="reports"></a><%end if%><BR>
  <a href="logincheck.asp?logoff=1" onMouseOver="act('logoff')" onMouseOut="inact('logoff')"><img border=0 src="img/menu_logoff.gif" name="logoff" id="logoff"></a><BR>
  <img src="img/spacer_blue2.gif" width=152 height=15><BR><BR>
    </TD></TR>
        <TR> 
          <TD><img src="img/menu.gif"><BR>
<%if go="h" then%><img border=0 src="img/menuon_xhome.gif"><%else%><a href="tool.asp?go=h" onMouseOver="act('xhome')" onMouseOut="inact('xhome')"><img border=0 src="img/menu_xhome.gif" name="xhome" id="xhome"></a><%end if%><BR>
<%if go="a" then%><img border=0 src="img/menuon_about.gif"><%else%><a href="tool.asp?go=a" onMouseOver="act('about')" onMouseOut="inact('about')"><img border=0 src="img/menu_about.gif" name="about" id="about"></a><%end if%><BR>
<%if go="y" then%><img border=0 src="img/menuon_xdemo.gif"><%else%><a href="tool.asp?go=y" onMouseOver="act('xdemo')" onMouseOut="inact('xdemo')"><img border=0 src="img/menu_xdemo.gif" name="xdemo" id="xdemo"></a><%end if%><BR>
<%if go="z" then%><a href="tool.asp?go=z"><img border=0 src="img/menuon_resources.gif"></a><%else%><a href="tool.asp?go=z" onMouseOver="act('resources')" onMouseOut="inact('resources')"><img border=0 src="img/menu_resources.gif" name="resources" id="resources"></a><%end if%><BR>
<%if go="c" then%><img border=0 src="img/menuon_contact.gif"><%else%><a href="tool.asp?go=c" onMouseOver="act('contact')" onMouseOut="inact('contact')"><img border=0 src="img/menu_contact.gif" name="contact" id="contact"></a><%end if%><BR>
		  <img src="img/spacer_blue2.gif" width=152 height=15><BR><BR></TD>
        </TR>
	
	
	</TBODY></TABLE><BR>
  </DIV>
<% end sub %>