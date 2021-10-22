<LINK href="img/weblog2.css" type=text/css rel=stylesheet>
<% sub questionaire 
sphere = cint(Request.QueryString("sphere"))
spheretext = Request.QueryString("t")
Dim s(20)


%>
<script>
function deactivate(number){
  document.getElementById('a' + number).bgColor = '#BBBBBB';
  }
function activate(number){
  document.getElementById('a' + number).bgColor = '#e47905';
}  

</script>

      <DIV class=floatcolmiddlehome> 
<img align="left" src="img/title_questionaire.gif"><a href="javascript:Submita('tool.asp?go=q&sphere=<%=Request.QueryString("sphere")%>&t=<%=Request.QueryString("t")%>')"><img border=0 align="right" alt="Click to save changes made to this page (Note: you need not use this button as changes are saved automatically when you switch pages)" src="img/button_save.gif"></a><BR>
<BR>
<TABLE cellSpacing=0 cellPadding=4 border=2 bgcolor="#FF6666" width="100%"><TBODY><TR><TD>
<font color="#FFFFFF"><SPAN class=hed12px><strong>Currently selected analysis: <%=Session("scenariostitle")%> (<%=Session("scenarioscountry")%>)</strong></SPAN></FONT>
</TD></TR></TBODY></TABLE><BR><BR>
<TABLE cellSpacing=0 cellPadding=0 width=590 border=0><TBODY>
	  <TR align=center>
	  <TD width=95><FONT size=-2>Legitimacy<BR>score -></FONT></TD>
		<%Set rs = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT spheres.spheresID, spheres.title, score.score, spheres.ordr as o " _
	 		& "FROM score INNER JOIN spheres ON score.spheresID = spheres.spheresID " _
	 		& "WHERE ((score.scenariosID=" & Session("scenariosID") &  ")) ORDER BY spheres.ordr;"
		rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
	  for i = 1 to 9%>
	  		<TD><FONT size=-2><%if rs("score")=-1 then%>NA<%else%><%=rs("score")%><%end if%></FONT></TD>
	  <%rs.movenext
	  next
	  rs.movefirst%>	  
	  </TR>
	  <TR align=center>
	  <TD><FONT size=-2>Sphere -></FONT></TD>
	  <%for i = 1 to 9%>
	  <% s(i) =rs("title")%>
	   <TD>
	  	<%if i = sphere then%><a href="javascript:Submita('tool.asp?go=q&sphere=<%=rs("spheresID")%>&t=<%=rs("title")%>')" ><img src="img/tabon_<%=i%>.gif" border=0 ></a><%else%><a href="javascript:Submita('tool.asp?go=q&sphere=<%=rs("spheresID")%>&t=<%=rs("title")%>')" onMouseOver="act('tab<%=i%>')" onMouseOut="inact('tab<%=i%>')"><img border=0 src="img/tab_<%=i%>.gif" name="tab<%=i%>" id="tab<%=i%>"></a><%end if%></TD><%
	  rs.movenext
	  next
	  rs.close
      set rs = nothing %>
	  </TR>
</TBODY></TABLE>
	  <TABLE cellSpacing=0 cellPadding=0 width=590 border=0 bgcolor="#339900"><TBODY>
	  <TR align=left><TD style="PADDING-RIGHT: 1px; PADDING-LEFT: 4px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px"><FONT size =-2 color="#FFFFFF"><img align="absmiddle" src="img/bullet_arrow_red2.gif">&nbsp;<STRONG><%=spheretext%></STRONG></FONT></TD></TR>
	</TBODY></TABLE>
  
<input type="hidden" name="sphere" value="<%=sphere%>">
<input type="hidden" name="t" value="<%=spheretext%>">
<input type="hidden" name="ggoto" value="default.asp">
          <TABLE cellSpacing=0 cellPadding=5 width=579 bordercolor="#0E4B78" border=0><TBODY>
			<TR><TD>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT questions.qtype as qtype, questions.questionsID as qID, questions.ordr as ordr1, questions.text, questions.helpinfo, questions.numcomments, questions.scale_low, questions.scale_high, data.rating, questions.importance " _
	& "FROM (questions INNER JOIN data ON questions.questionsID = data.questionsID) " _
	& "WHERE (((data.scenariosID)=" & Session("scenariosID") & ") AND ((questions.spheresID)=" & sphere & ")) " _
	& "ORDER BY questions.ordr;"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
while rs.EOF = false
%>
          <TABLE cellSpacing=0 cellPadding=0 width=575 bgColor=#e47905 border=0 align=center><TBODY>
        <tr><td valign="top" bgcolor="#e47905" width="575" height="1"><font size="1"><img src="/wb/images/spacer.gif" height="1" width="520"></font></td></tr>

		  	<TR><TD>
			<TABLE cellSpacing=0 cellPadding=0 width=573 align=center  border=0><TBODY>
			<TR <%if rs("rating") = 0 then%>bgColor="#BBBBBB"<%else%>bgColor="#e47905"<%end if%> id=a<%=rs("ordr1")%> >
			<TD width=200 >&nbsp;&nbsp;<FONT size=-2 color="#FFFFFF">Question</FONT>&nbsp;&nbsp;<font size=+1 color="#FFFFFF"><strong><%=rs("ordr1")%></strong></font>&nbsp;&nbsp;
			<%select case rs("importance")
			Case "1" %>
			<img src="img/icon_l.gif" alt="LOW IMPORTANCE: This question has been given a low importance rating">
			<%Case "2" %>
			<img src="img/icon_m.gif" alt="MEDIUM IMPORTANCE: This question has been given a medium importance rating">
			<%Case "3"%>
			<img src="img/icon_h.gif" alt="HIGH IMPORTANCE: This question has been given a high importance rating">
			<%end Select%>
			&nbsp;
			<%select case rs("qtype")
			Case "2" %>
			<img src="img/icon_qtype_1.gif" alt="RULE: This question pertains to a rule for the sphere">
			<%Case "1" %>
			<img src="img/icon_qtype_2.gif" alt="META-RULE: This question pertains to a meta-rule for the sphere">
			<%Case "3"%>
			<img src="img/icon_qtype_3.gif" alt="EXOGENOUS FACTORS: This question pertains to exogenous factors for the sphere">
			<%end Select%>			
			
			</TD>
			<TD width=250 ><div align="left" id=b<%=rs("ordr1")%>><%if rs("rating") = 0 then%><FONT size=-2 color="#FFFFFF">&nbsp;<%end if%></div></TD>
			<TD width=125 align=right><%if len(rs("helpinfo"))>0 then%><a href="moreinfo.asp?qID=<%=rs("qID")%>" target="_blank"><img border=0 src="img/info_icon.gif" alt="INFORMATION: Click here for help in answering this question"></a><%else%><img src="img/spacer_trans.gif" height="15" width="15"><%end if%>&nbsp;
			<a href="comments.asp?qID=<%=rs("qID")%>&numcomments=<%=rs("numcomments")%>" target="_blank"><img border=0 src="img/memo_icon.gif" alt="USER COMMENTS: Click here to view or add"><FONT size=-2><%=rs("numcomments")%></FONT></a>&nbsp;&nbsp;</TD>
			</TR></TBODY></TABLE>
			</TD></TR>
			</TBODY></TABLE>
			<TABLE cellSpacing=1 cellPadding=8 width=575 bgColor=#e47905 border=0 align=center><TBODY>
              <TR> 
				<TD valign="top" bgcolor="#EEEEEE"><font size="-1"><%=rs("text")%></FONT>
				<BR>
				</TD>
			  </TR>
			  <TR>
			    <TD bgcolor="#FCE2C7" align="center">
				<TABLE><TBODY>
				<TR align="center">
				<TD>&nbsp;</TD>
				<TD><FONT size=-2><strong>N/A&nbsp;&nbsp;&nbsp;</strong></FONT></TD><TD>&nbsp;</TD><%for i = 1 to 7%><TD><FONT size=-2><strong><%=i%></strong></FONT></TD><%next%><TD>&nbsp;</TD>
				</TR>
				<TR align="center">
				<TD><FONT size=-1><strong>RATING:&nbsp;&nbsp;</strong></FONT></TD>
				<TD><INPUT TYPE=RADIO onClick="javascript:deactivate('<%=rs("ordr1")%>')" NAME="rating_<%=rs("ordr1")%>" VALUE="0" <%if rs("rating") = 0 then%> CHECKED <%end if%>>&nbsp;|</TD>
				<TD><%if len(rs("scale_low"))>0 then %><FONT size=-2><strong><%=rs("scale_low")%></strong></FONT><%else%>&nbsp;<%end if%></TD>
				<%for i = 1 to 7%><TD><INPUT TYPE=RADIO onClick="javascript:activate('<%=rs("ordr1")%>')" NAME="rating_<%=rs("ordr1")%>" VALUE="<%=i%>" <%if rs("rating") = i then%> CHECKED <%end if%>></TD><%next%>
				<TD><%if len(rs("scale_high"))>0 then %><FONT size=-2><strong><%=rs("scale_high")%></strong></FONT><%else%>&nbsp;<%end if%></TD>
				</TR>
				</TBODY></TABLE>
				</TD>
			  </TR>			  
            </TBODY>
          </TABLE>
<BR>
<%
rs.movenext
	wend
	rs.close
	set rs = nothing
%>			  
		</TD></TR></TBODY></TABLE>
</FORM>
<% if sphere = "9" then%>
	  <a href="javascript:Submita('tool.asp?go=e')"><img align="right" src="img/button_next.gif" border=0 alt="STEP 3: Results"></a>
<% else %>
	  <a href="javascript:Submita('tool.asp?go=q&sphere=<%=sphere+1%>&t=<%=s(sphere+1)%>')" ><img align="right" src="img/button_nextsphere.gif" border=0 alt="Next Sphere: <%=s(sphere+1)%>"></a>
<% end if%>
      </DIV>

<% end sub %>