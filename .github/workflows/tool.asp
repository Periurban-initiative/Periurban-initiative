<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML><HEAD><TITLE>DSTAIR: A Decision Support Tool to Analyze Institutional Reform</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<META 
content="DSTAIR, Decision Support Tool, Institutional Reform, corruption" 
name=keywords>
<META 
content="DSTAIR: A Decision Support Tool to Analyse Institutional Reform" 
name=description>
<LINK href="img/weblog2.css" type=text/css rel=stylesheet>
<LINK href="img/horizontalmenus.css" type=text/css rel=stylesheet>
<LINK href="images/urlicon.ico" rel="shortcut icon">
<META content="MSHTML 6.00.2800.1476" name=GENERATOR></HEAD>
<BODY>
<!--#include file ="getFileContents.asp"-->
<!--#include file ="login.asp"-->
<!--#include file ="conn.inc"-->
<!--#include file ="menu.asp"-->
<!--#include file ="questionaire.asp"-->
<!--#include file ="results.asp"-->
<!--#include file ="welcome.asp"-->
<!--#include file ="analysis.asp"-->
<!--#include file ="actools.asp"-->
<!--#include file ="progressbar.js"-->
<SCRIPT type=text/javascript>
function Submita(ggoto) {
<% if Request.QueryString("go") = "q" then %>
	document.forms['data'].ggoto.value = ggoto;
	scroll(0,0);
	CallJS('Demo()');
	document.forms['data'].submit();
<% else %>
	window.location=ggoto;
<% end if%>
}
function Submitaa(ggoto) {
	document.forms['changeanalysis'].submit();
}

if (window.createPopup && document.compatMode && document.compatMode=="CSS1Compat")
{
  document.onreadystatechange = onresize = function fixIE6AbsPos()
  {
    if (!document.body) return;
    if (document.body.style.margin != "0px") document.body.style.margin = 0;
    onresize = null;
    document.body.style.height = 0;
    setTimeout(function(){ document.body.style.height = document.documentElement.scrollHeight+'px'; }, 1);
    setTimeout(function(){ onresize = fixIE6AbsPos; }, 100);
  }
}
</SCRIPT>

<DIV id=pagewrapper> 
  <%		
usersID = Session("usersID")
go = Request.QueryString("go")
		if (usersID = -1) or  (len(usersID) = 0) then
			Response.Redirect ("default.asp")
		else if Session("usersID") = -2 then
			admin
		else
			menu
			%>
  <DIV id=contentouter> 
    <DIV id=contentcontainer> 
      <DIV id=content> 
        <%
			select case go
			Case "i"
				getfilecontents("instructions")
			case "s"
				select case Request.QueryString("function") 
				Case "edit"
					editanalysis
				Case "add"
					addanalysis
				Case Else
					analysis
				end select
			Case "q"
				questionaire
			Case "e"
				results
			Case "t"
				actools
			Case "r"
				getfilecontents("reports")
			Case "h"	
				getfilecontents("home")
			Case "a"
				getfilecontents("about")
			Case "y"
				getfilecontents("instructions")
			Case "z"
				getfilecontents("resources")

			Case "r1"
				getfilecontents("toolkits")
			Case "r2"
				getfilecontents("faqs")
			Case "c"
				getfilecontents("contact")
			Case else
				welcome
			end select
			%>
        <DIV class=tinytext align=center> 
          <table width="540" border="0" cellspacing="0" cellpadding="0" align="center">
            <tr> 
              <td valign="top" width="740" height="4"><font size="1"><img src="/wb/images/spacer.gif" height="4" width="520"></font></td>
            </tr>
            <tr> 
              <td valign="top" bgcolor="#000000" width="740" height="1"><font size="1"><img src="/wb/images/spacer.gif" height="1" width="520"></font></td>
            </tr>
            <tr> <BR>
              <td valign="top" width="740" height="4"><font size="1"><img src="/wb/images/spacer.gif" height="4" width="520"></font></td>
            </tr>
            <tr> 
              <td>Copyright © 2005 <a href="http://www.tellus.org" target="_blank"><font color="#0033CC">Tellus 
                Institute</font></a> | All Rights Reserved</td>
            </tr>
          </table>
          <BR>
        </DIV>
      </DIV>
    </DIV>
  </DIV>
  <%
		end if
		end if
%>
  <DIV id=hedlogo><A href="http://www.institutionalreform.org" target=_top><IMG 
height=77 alt="DSTAIR" 
src="img/homelogo2.gif" border=0></A></DIV>
  <DIV id=hedpict><img height=77 alt="DSTAIR" 
src="img/header.gif" width=600 border=0 
name=image> </DIV>
</DIV>
	  <SCRIPT LANGUAGE="JavaScript">

// Create layer for progress dialog
document.write("<span id=\"progress\" class=\"hide\">");
	document.write("<FORM name=dialog>");
	document.write("<TABLE border=2  bgcolor=\"#FF0000\">");
	document.write("<TR><TD ALIGN=\"center\">");
	document.write("<FONT color=\"#FFFFFF\"><STRONG>Saving ...<BR></STRONG></FONT>");
	document.write("<input type=text name=\"bar\" size=\"" + _progressWidth/2 + "\"");
	if(document.all||document.getElementById) 	// Microsoft, NS6
		document.write(" bar.style=\"color:navy;\">");
	else	// Netscape
		document.write(">");
	document.write("</TD></TR>");
	document.write("</TABLE>");
	document.write("</FORM>");
document.write("</span>");
ProgressDestroy();	// Hides

</script>
</BODY></HTML>
