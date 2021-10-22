<!--#include file ="conn.inc"-->
<!--#include file ="programs.asp"-->

<%
Response.Buffer = true
if len(Request.QueryString("logoff"))>0 then
	Response.Clear
	Session.Abandon
	response.redirect("default.asp")
else
	Set rs = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM users WHERE userid='" & Request.Form("userid") & "';"
	rs.Open strSQL, conn, 1,3
	if rs.EOF = true then
		rs.close
		set rs = nothing
		Response.Clear
		response.redirect("default.asp?error=1")
	else
		if rs("pw")=Request.Form("pw") then
			if Request.Form("userid") = "admin" then
				Session("usersID") = -2	
			else
				Session("usersID") = rs("usersID")
			end if	
			Session("Userid") = rs("userid")
			Session("Name") = rs("name")
			Session.Timeout=60
			if isnull(rs("lastdate")) then
				session("new") = 1
			end if
			Session("Lastlogin") = rs("lastdate")
			rs("lastdate") = now()
			rs.update
			rs.close
			set rs = nothing
			if Session("usersID") >0 then
				'get first scenario
				Set rs = Server.CreateObject("ADODB.Recordset")
				strSQL = "SELECT * FROM scenarios WHERE usersID=" & Session("usersID") & " ORDER BY scenariosID;"
				rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
				Session("scenariosID") = rs("scenariosID")
				Session("scenariostitle") = rs("title")
				Session("scenarioscountry") = rs("country")
				rs.close
				set rs = nothing
			end if
			'count spheres
			Set rs = Server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT count(*) as x FROM spheres;"
			rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
			Session("numspheres") = rs("x")
			rs.close
			set rs = nothing			
			'count tools
			Set rs = Server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT count(*) as x FROM tools;"
			rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
			Session("numtools") = rs("x")
			rs.close
			set rs = nothing			
			Response.Clear
			if Request.Form("userid") = "reset" then
				update_dataset
				response.redirect("logincheck.asp?logoff=1")
			else
				response.redirect("tool.asp")
			end if
		else
			rs.close
			set rs = nothing		
			Response.Clear
			response.redirect("default.asp?error=1")
		end if
	end if
end if		
%>


<% sub checkscenario 
'this subroutine checks to see if this user has logged in before and whether there
'are records associated with their scenario, and generates the records if necessary
Set rs = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM data WHERE scenariosID =" & Session("scenariosID") & ";"
rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
if rs.eof = true then
	rs.close
	set rs = nothing
	'DATA: open recordset of default scenario (scenarioID=0)
	Set rs = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM data WHERE scenariosID = 0 ORDER BY questionsID;"
	rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
	'DATA: open recordset for writing new records
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM data;"
	rs2.Open strSQL, conn, 1,3
	'copy default values to new scenario
	while rs.EOF = false
		rs2.addnew
		rs2("scenariosID") = cint(Session("scenariosID"))
		rs2("questionsID") = rs("questionsID")
		rs2("rating") = rs("rating")
		rs2("importance") = rs("importance")
		rs2.update
		rs.movenext
	wend
	rs.close
	set rs = nothing
	rs2.close
	set rs2 = nothing
	'INFLUENCES: open recordset of default scenario (scenarioID=0)
	Set rs = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT influences.* FROM data  " _
	& "INNER JOIN influences ON data.dataID = influences.dataID " _
	& "WHERE (((data.scenariosID)=0)) ORDER BY influences.dataID,influences.spheresID;"
	rs.Open strSQL, conn, adOpenForwardOnly, , adCmdText
	'INFLUENCES: open recordset for writing new records
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM influences;"
	rs2.Open strSQL, conn, 1,3
	'copy default values to new scenario
	while rs.EOF = false
		rs2.addnew
		rs2("dataID") = rs("dataID")
		rs2("spheresID") = rs("spheresID")
		rs2("value") = rs("value")
		rs2.update
		rs.movenext
	wend
	rs.close
	set rs = nothing
	rs2.close
	set rs2 = nothing
end if
end sub
%>

