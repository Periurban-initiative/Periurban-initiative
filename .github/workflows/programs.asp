	<%
sub calculate(spheresID)

if spheresID<0 then
	spheresID = Request.Form("sphere")
end if	

' CALCULATE A: internal trustworthiness for each sphere
dim A(10)
dim A1
dim A2
for j = 1 to Session("numspheres")
	Set rs = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT questions.spheresID AS j, data.scenariosID, data.rating as rating, questions.importance as importance " _
	& "FROM (questions INNER JOIN data ON questions.questionsID = data.questionsID) " _
	& "WHERE ((questions.spheresID)=" & j & ") AND ((data.scenariosID)=" & Session("scenariosID") & ") " _
	& "AND ((data.rating)>0);"
	rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
	if rs.EOF then
		A(j) = -1
	else
		A1= 0
		A2= 0
		i = 0
		while rs.EOF = false
			i = i + 1
			A1 = A1 + ((rs("rating")) * rs("importance"))
			A2 = A2 + rs("importance")
			rs.movenext
		wend
		A(j) = A1/(7 * A2)
	end if	
	rs.close
	set rs = nothing
next

'save calculations to database
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM score WHERE scenariosID=" & Session("scenariosID") & " ORDER BY spheresID;"
rs.Open strSQL, conn, 1,3
for j = 1 to Session("numspheres")
	rs("score") = round(A(j),3)
	rs.update
	rs.movenext
next	
rs.close
set rs = nothing

'update tools
calc_tools(session("scenariosID"))

end sub

sub savedata
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT questions.ordr as ordr1, data.rating as rating " _
	& "FROM (questions INNER JOIN data ON questions.questionsID = data.questionsID) " _
	& "WHERE (( (data.scenariosID)=" & Session("scenariosID") & ") AND ((questions.spheresID)=" & cstr(Request.Form("sphere")) & ")) " _
	& "ORDER BY questions.ordr;"
rs.Open strSQL, conn, 1,3
while rs.EOF = false
	rs("rating") = cint(Request.Form("rating_" & rs("ordr1")))
	rs.update
	rs.movenext
wend
rs.close
set rs = nothing

Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "select * from scenarios WHERE scenariosID=" & Request.Form("analysis") & ";"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
if rs.EOF = FALSE then
	session("scenariosID") = cint(Request.Form("analysis"))
	Session("scenariostitle") = rs("title")
	Session("scenarioscountry") = rs("country")
end if
rs.close
set rs = nothing

end sub


sub update_dataset
' delete all scenarios
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "select * from scenarios;"
rs.Open strSQL, conn, 1,3
while rs.EOF = false
	rs.delete
	rs.movenext
wend	
rs.close
set rs = nothing

' delete all data 
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "select * from data;"
rs.Open strSQL, conn, 1,3
while rs.EOF = false
	rs.delete
	rs.movenext
wend	
rs.close
set rs = nothing

' delete all scores
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM score;"
rs.Open strSQL, conn, 1,3
while rs.EOF = false
	rs.delete
	rs.movenext
wend
rs.close
set rs = nothing

' delete all tooldata
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM tooldata;"
rs.Open strSQL, conn, 1,3
while rs.EOF = false
	rs.delete
	rs.movenext
wend
rs.close
set rs = nothing

'add scenario scenario for each user already in the system
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM users WHERE usersID > 0;"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
while rs.EOF = false
	add_scenario(rs("usersID"))
	rs.movenext
wend
rs.close
set rs = nothing

end sub 'reset (update dataset)

sub add_scenario(usersID)
'add scenario to list
dim scenID
Set rs4 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM scenarios;"
rs4.Open strSQL, conn, 1,3
rs4.addnew
rs4("usersID") = usersID
rs4("title") = "Default"
rs4("share")=false
rs4.update
scenID = rs4("scenariosID")
rs4.close
set rs4 = nothing

Set rs2 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM data;"
rs2.Open strSQL, conn, 1,3
Set rs3 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM questions;"
rs3.Open strSQL, conn, adOpenForwardOnly, , adCmdText
while rs3.EOF = false
	rs2.addnew
	rs2("scenariosID") = scenID
	rs2("questionsID") = rs3("questionsID")
	rs2("rating") = 0
	rs2.update
	rs3.movenext
wend
rs2.close
set rs2 = nothing
rs3.close
set rs3 = nothing

'add empty score records for new scenario
Set rs2 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM score;"
rs2.Open strSQL, conn, 1,3
for i = 1 to Session("numspheres")
	rs2.addnew
	rs2("scenariosID") = scenID
	rs2("spheresID") = i
	rs2("score") = -1
	rs2.update
next
rs2.close
set rs2 = nothing

'add empty tooldata records for new scenario
Set rs2 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM tooldata;"
rs2.Open strSQL, conn, 1,3
for i = 1 to Session("numtools")
	rs2.addnew
	rs2("scenariosID") = scenID
	rs2("toolID") = i
	rs2("trigger") = Session("numspheres")
	rs2.update
next
rs2.close
set rs2 = nothing
end sub 'add_scenario
	
sub addnewuser
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT max(usersID) as m FROM scenarios;"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
usersID = rs("m") + 1

Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM users;"
rs.Open strSQL, conn, 1,3
	rs.addnew
	rs("usersID") = usersID
	rs("userID") = Request.Form("userid")
	rs("pw") = Request.Form("pw")
	rs("name") = Request.Form("name")
	rs("lastdate") = null
	rs.update
rs.close
set rs = nothing
add_scenario(cint(usersID))
end sub

sub calc_tools(scenID)

Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM score WHERE scenariosID = " & scenID & " ORDER BY spheresID;"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
Set rs2 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM tooldata WHERE scenariosID = " & scenID & ";"
rs2.Open strSQL, conn, 1,3
while rs2.EOF = false
	Set rs3 = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM toolcriteria WHERE toolID = " & rs2("toolID") & " ORDER BY spheresID;"
	rs3.Open strSQL, conn, adOpenForwardOnly, , adCmdText
	x = 0
	rs.movefirst
	while rs.EOF = false
		if (rs("score") >= rs3("data")) or (rs("score")=-1) then
			x = x + 1
		end if
		rs.movenext
		rs3.movenext
	wend
	rs2("trigger") = x
	rs2.update		
	rs2.movenext
	rs3.close
	set rs3 = nothing
wend
rs.close
set rs = nothing
rs2.close
set rs2 = nothing

end sub

sub create_empty_toolcriteria
Set rs2 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM toolcriteria;"
rs2.Open strSQL, conn, 1,3
for i = 1 to cint(Session("numtools"))
for j = 1 to cint(Session("numspheres"))
	rs2.addnew
	rs2("toolID")=i
	rs2("spheresID")=j
	rs2("data")=0
	rs2.update
next
next
rs2.close
set rs2 = nothing

end sub


sub savetoolscriteria
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT toolID, data, spheresID FROM toolcriteria " _
	& "ORDER BY toolID,spheresID;"
rs.Open strSQL, conn, 1,3
while rs.EOF = false
	for j = 1 to 9
		z = Request.Form("tool_" & rs("toolID") & "_" & j)
		if (len(z) = 0) or (isnumeric(z) = false) then
			z = 0
		end if
		rs("data") = cdbl(z)
		rs.update
		rs.movenext
	next		
wend
rs.close
set rs = nothing
'recalc for each scenario
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM scenarios;"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
while rs.EOF = false
	calc_tools(rs("scenariosID"))
	rs.movenext
wend
rs.close
set rs = nothing

end sub

sub yuck
%>
<a href="editcomment.asp?commentID=<%=rs("commentsID")%>"><img src="img/edit_icon.gif" title="Edit comment"></a>
<%end sub
%>