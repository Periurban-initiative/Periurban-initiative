<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">

<HTML><HEAD><TITLE>DSTAIR: A Decision Support Tool to Analyze Institutional Reform</TITLE><!-- Load the header, including menu scripts-->
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<META 
content="DSTAIR, Decision Support Tool, Institutional Reform, corruption" 
name=keywords>
<META 
content="DSTAIR: A Decision Support Tool to Analyze Institutional Reform" 
name=description>
<LINK href="img/weblog2.css" type=text/css rel=stylesheet>
<LINK href="img/horizontalmenus.css" type=text/css rel=stylesheet>
<LINK href="images/urlicon.ico" rel="shortcut icon">
<META content="MSHTML 6.00.2800.1476" name=GENERATOR></HEAD>
<BODY>
<!-- Start of StatCounter Code -->
<script type="text/javascript" language="javascript">
var sc_project=823962; 
var sc_partition=6; 
var sc_security="ad59f9bc"; 
</script>

<script type="text/javascript" language="javascript" src="http://www.statcounter.com/counter/counter.js"></script><noscript><a href="http://www.statcounter.com/" target="_blank"><img  src="http://c7.statcounter.com/counter.php?sc_project=823962&amp;java=0&amp;security=ad59f9bc" alt="web stats analysis" border="0"></a> </noscript>
<!-- End of StatCounter Code -->
<!--#include file ="getFileContents.asp"-->
<SCRIPT language=JavaScript>
<!--
 
	if (document.images) {
<% For Each x in array("xhome", "about", "xinfo", "xdemo", "resources", "contact", "access", "logoff") %>
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

<SCRIPT type=text/javascript>
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
<%go2 = Request.QueryString("go")
if len(go2)=0 then
	go2 = "h"
end if
go = left(go2,1)
%>

<DIV id=pagewrapper> 
  <DIV id=leftcolumn> 
    <TABLE width=150 cellspacing="0" border="0" cellpadding="0">
      <TBODY>
        <TR> 
          <TD><img src="img/menu.gif"><BR>
<%if go="h" then%><img border=0 src="img/menuon_xhome.gif"><%else%><a href="default.asp?go=h" onMouseOver="act('xhome')" onMouseOut="inact('xhome')"><img border=0 src="img/menu_xhome.gif" name="xhome" id="xhome"></a><%end if%><BR>
<%if go="i" then%><img border=0 src="img/menuon_about.gif"><%else%><a href="default.asp?go=i" onMouseOver="act('about')" onMouseOut="inact('about')"><img border=0 src="img/menu_about.gif" name="about" id="about"></a><%end if%><BR>
<%if go="d" then%><img border=0 src="img/menuon_xdemo.gif"><%else%><a href="default.asp?go=d" onMouseOver="act('xdemo')" onMouseOut="inact('xdemo')"><img border=0 src="img/menu_xdemo.gif" name="xdemo" id="xdemo"></a><%end if%><BR>
<%if go="r" then%><a href="default.asp?go=r"><img border=0 src="img/menuon_resources.gif"></a><%else%><a href="default.asp?go=r" onMouseOver="act('resources')" onMouseOut="inact('resources')"><img border=0 src="img/menu_resources.gif" name="resources" id="resources"></a><%end if%><BR>
<%if go="c" then%><img border=0 src="img/menuon_contact.gif"><%else%><a href="default.asp?go=c" onMouseOver="act('contact')" onMouseOut="inact('contact')"><img border=0 src="img/menu_contact.gif" name="contact" id="contact"></a><%end if%><BR>
		  <img src="img/spacer_blue2.gif" width=152 height=15><BR><BR></TD>
        </TR>
        <TR> 
          <TD style="PADDING-RIGHT: 1px; PADDING-LEFT: 1px; PADDING-BOTTOM: 1px; PADDING-TOP: 1px" bgColor=#e9f4ff> 
            <IMG alt="" src="img/menu1.gif"> 
              <%if request.QueryString("error")=1 then%><DIV class=hed10px>
              <strong><FONT color="#FF0000">Incorrect login</FONT></strong></DIV><BR>
              <%end if%>
			  <%if len(session("scenariosID")) > 0 then %>
			  <TABLE bordercolor="#6699CC" border=1 width=152 cellpadding="0" cellspacing="0"><TBODY><TR><TD><BR><DIV class=hed10px align=center><strong>You are logged in as</strong><BR>
			  <%=Session("Name")%></DIV><BR></TD></TR></TBODY></TABLE>
			  <a href="tool.asp" onMouseOver="act('access')" onMouseOut="inact('access')"><img border=0 src="img/menu_access.gif" name="access" id="access"></a><BR>
			  <a href="logincheck.asp?logoff=1" onMouseOver="act('logoff')" onMouseOut="inact('logoff')"><img border=0 src="img/menu_logoff.gif" name="logoff" id="logoff"></a><BR>


			  <%else%>
			  <form action="logincheck.asp" method="post" name="data">
              <DIV class=hed10px><strong>USERID:</strong><BR>
                <input type="text" name="userid">
              </div>
              <DIV class=hed10px><strong>PASSWORD:</strong><BR>
                <input type="password" name="pw">
              </div>
              <INPUT type="submit" value="SUBMIT">
              </form>
			  <%end if%>
            <BR></TD>
        </TR>
      </TBODY>
    </TABLE>
    <BR>
  </DIV>
  <DIV id=contentouter> 
  <DIV id=contentcontainer> 
  <DIV id=content> 

    <%select case go2
case "h": getfilecontents("home")
case "i": getfilecontents("about")
case "i2": getfilecontents("moreinfo")
case "c": getfilecontents("contact")
case "d": getfilecontents("howitworks")
case "r": getfilecontents("resources")
case "r1": getfilecontents("toolkits")
case "r2": getfilecontents("faqs")
end select
%>
  </DIV></DIV>
  <td>Copyright © 2005 <a href="http://www.tellus.org" target="_blank"><font color="#0033CC">Tellus 
                Institute</font></a> | All Rights Reserved</td>
</DIV>
  <DIV id=hedlogo><IMG height=77 alt="DSTAIR" src="img/homelogo2.gif" width=160 border=0></A></DIV>
  <DIV id=hedpict><img height=77 width=600 alt="DSTAIR" src="img/header.gif"  border=0 name=image> </DIV>
</DIV>
</BODY></HTML>
