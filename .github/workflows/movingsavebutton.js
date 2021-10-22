
var ns = (document.layers)?1:0;

/*
Button:
	this code is to show a "button" to switch
	the fly ON/OFF.
	It is always shown on the frame's
	bottom-right corner.
------------------------------------------------
*/
var imgwidth=40;	// Image width
var imgheight=40;	// Image height

var button = Array();	// to pre-cache images
button[0]=new Image();
button[0].src="img/info_icon.gif";
button[1]=new Image();
button[1].src="img/info_icon.gif";

var text="<table width=10 bgcolor=#ffffff><td><a href='google.com'><center><img name='Button' src='"+button[0].src+"' width='"+imgwidth+"' height='"+imgheight+"' border='0'></center></a></font></td></table>"	// A bit of HTML code to display the button


//Initialize some variables to make the button always to appear on the frame's bottom-right corner
if (ns) {
	document.write("<LAYER NAME='FlyOnOff' LEFT=0 TOP=0>"+text+"</LAYER>");
	horz=".left";
	vert=".top";
	docStyle="document.";
	styleDoc="";
	innerW="window.innerWidth";
	innerH="window.innerHeight";
	offsetX="window.pageXOffset";
	offsetY="window.pageYOffset";
}else {
	document.write("<div id='FlyOnOff' style='position:absolute; visibility:show; left:235px; top:-50px; z-index:2'>"+text+"</div>");
	horz=".pixelLeft";
	vert=".pixelTop";
	docStyle="";
	styleDoc=".style";
	innerW="document.body.clientWidth";
	innerH="document.body.clientHeight";
	offsetX="document.body.scrollLeft";
	offsetY="document.body.scrollTop";
}


// Moves the button in the right position
function checkLocation() {
	objectXY="FlyOnOff";
	var availableX=eval(innerW);
	var availableY=eval(innerH);
	var currentX=eval(offsetX);
	var currentY=eval(offsetY);
	x=availableX-(imgwidth+30)+currentX;
	y=availableY-(imgheight+20)+currentY;
	eval(docStyle + objectXY + styleDoc + horz + "=" + x);
	eval(docStyle + objectXY + styleDoc + vert + "=" + y);
}

setInterval('checkLocation()', 10);

/*
end of Button management
------------------------------------------------
*/
