<%

sub admin
if Session("usersID") <> -2 then
	Response.Clear
	Session.Abandon
	response.redirect("default.asp")
end if
%>
<DIV id=pagewrapper> 
  <DIV id=contentouter> 
    <DIV id=contentcontainer> 
      <DIV id=content> 

    <DIV class=floatcolmiddlehome> 

<TABLE cellSpacing=0 cellPadding=6 width="100%" border=0><TBODY>
<TR> 
	<TD class=sidebar1 id=middleblocks vAlign=top width=100% background="img/bg_grid.gif" name="middleblocks">
		<table width=85% cellspacing=1 cellpadding=6 border=1><tbody>
		<tr bgcolor="#FFFF66"><td width="10%">ID</td>
		<td width="10%">USERID/<BR>PASSWORD</td>
		<td width="25%">LAST<BR>LOGIN</td>
		<td width="25%">NAME</td>
		</tr>
<%
	Set rs = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM users ORDER BY usersID;"
	rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
	while rs.EOF=false 
%>
		<tr><td width="10%"><%=rs("usersID")%></td>
		<td width="25%"><%=rs("userid")%><BR><%=rs("pw")%></td>
		<td width="25%"><%=rs("lastdate")%></td>
		<td width="30%"><%=rs("name")%></td>
		</tr>				
<% rs.movenext
wend 
		rs.close
		set rs = nothing
%>
		</table>
		<br><br><strong><font size="+2">Add New User</font></strong>
		<form action="addnewuser.asp" method="post" name="data">
		<table width=50% cellspacing=1 cellpadding=6 border=1><tbody>
		<tr><td width=25% align="right"><strong>USERID:</strong></td><td width=75% ><input type="text" name="userid"></td></tr>
		<tr><td width=25% align="right"><strong>PASSWORD:</strong></td><td width=75% ><input type="password" name="pw"></td></tr>
		<tr><td width=50% align="right"><strong>NAME:</strong></td><td width=75% ><input type="name" name="name"></td></tr>		
		</tbody></table><br>
		<INPUT type="submit" value="SUBMIT"> 
		</form><BR><BR>
		<a style="color:#330099" href="logincheck.asp?logoff=1"><FONT size=+2><strong>LOGOFF</strong></FONT></a>
	</TD>
</TR>
</TBODY></TABLE>
</DIV>
</DIV>
</DIV>
</DIV>
</DIV>


<%end sub%>