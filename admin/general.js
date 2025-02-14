function clearHTMLChar(str)
{
	var retVal = str;
	var arrStrRep = new Array	(
									new Array('&#225;', 'á'),
									new Array('&#233;', 'é'),
									new Array('&#237;', 'í'),
									new Array('&#243;', 'ó'),
									new Array('&#250;', 'ú'),
									new Array('&#241;', 'ñ'),
									new Array('&#191;', '¿'),
									new Array('&#193;', 'Á'),
									new Array('&#201;', 'É'),
									new Array('&#205;', 'Í'),
									new Array('&#211;', 'Ó'),
									new Array('&#218;', 'Ú'),
									new Array('&#209;', 'Ñ'),
									new Array('&#220;', 'Ü'),
									new Array('&#225;', 'á'),
									new Array('&#233;', 'é'),
									new Array('&#237;', 'í'),
									new Array('&#243;', 'ó'),
									new Array('&#250;', 'ú'),
									new Array('&#241;', 'ñ'),
									new Array('&#252;', 'ü'),
									new Array('&#192;', 'À'),
									new Array('&#194;', 'Â'),
									new Array('&#195;', 'Ã'),
									new Array('&#202;', 'Ê'),
									new Array('&#212;', 'Ô'),
									new Array('&#213;', 'Õ'),
									new Array('&#224;', 'à'),
									new Array('&#226;', 'â'),
									new Array('&#227;', 'ã'),
									new Array('&#234;', 'ê'),
									new Array('&#232;', 'è'),
									new Array('&#244;', 'ô'),
									new Array('&#245;', 'õ'),
									new Array('&#252;', 'ü'),
									new Array('&#251;', 'û'),
									new Array('&#199;', 'Ç'),
									new Array('&#231;', 'ç'),
									new Array('&quot;', '\"'),
									new Array('&amp;', '&'),
									new Array('&#161;', '¡'),
									new Array('&#1488;', 'א'),
									new Array('&#1489;', 'ב'),
									new Array('&#1490;', 'ג'),
									new Array('&#1491;', 'ד'),
									new Array('&#1492;', 'ה'),
									new Array('&#1493;', 'ו'),
									new Array('&#1494;', 'ז'),
									new Array('&#1495;', 'ח'),
									new Array('&#1496;', 'ט'),
									new Array('&#1497;', 'י'),
									new Array('&#1498;', 'ך'),
									new Array('&#1499;', 'כ'),
									new Array('&#1500;', 'ל'),
									new Array('&#1501;', 'ם'),
									new Array('&#1502;', 'מ'),
									new Array('&#1503;', 'ן'),
									new Array('&#1504;', 'נ'),
									new Array('&#1505;', 'ס'),
									new Array('&#1506;', 'ע'),
									new Array('&#1507;', 'ף'),
									new Array('&#1508;', 'פ'),
									new Array('&#1509;', 'ץ'),
									new Array('&#1510;', 'צ'),
									new Array('&#1511;', 'ק'),
									new Array('&#1512;', 'ר'),
									new Array('&#1513;', 'ש'),
									new Array('&#1514;', 'ת'),
									new Array('&#1523;', '׳'),
									new Array('&#1524;', '״'),
									new Array('&#8362;', '₪'),
									new Array('&#169;', '©'),
									new Array('&#174;', '®'),
									new Array('&#8217;', '\'')
								);
	for (var s = 0;s<arrStrRep.length;s++)
	{
		while (str.indexOf(arrStrRep[s][0]) != -1)
			str = str.replace(arrStrRep[s][0], arrStrRep[s][1]);
	}
	return str;
}

function chkWin() { if (OpenWin != null) if (!OpenWin.closed) OpenWin.focus() }
function clearWin() { OpenWin = null; }

function browserDetect(){
	detect = navigator.userAgent.toLowerCase();
    if(detect.indexOf('msie')+1){    
        return 'msie'; 
    }
    else if(detect.indexOf('firefox')+1){    
        return 'firefox';
    }
    else if (detect.indexOf('safari')+1){
    	return 'safari';
    }
    else if (detect.indexOf('opera')+1){
    	return 'opera';
    }
    else {
        return'firefox';
    }
}

function emailCheck (emailStr) {

var checkTLD=1;
var knownDomsPat=/^(com|net|org|edu|int|mil|gov|arpa|biz|aero|name|coop|info|pro|museum)$/;
var emailPat=/^(.+)@(.+)$/;
var specialChars="\\(\\)><@,;:\\\\\\\"\\.\\[\\]";
var validChars="\[^\\s" + specialChars + "\]";
var quotedUser="(\"[^\"]*\")";
var ipDomainPat=/^\[(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})\]$/;
var atom=validChars + '+';
var word="(" + atom + "|" + quotedUser + ")";
var userPat=new RegExp("^" + word + "(\\." + word + ")*$");
var domainPat=new RegExp("^" + atom + "(\\." + atom +")*$");

var matchArray=emailStr.match(emailPat);
if (matchArray==null) {
return false;
}
var user=matchArray[1];
var domain=matchArray[2];

for (i=0; i<user.length; i++) {
if (user.charCodeAt(i)>127) {
return false;
   }
}
for (i=0; i<domain.length; i++) {
if (domain.charCodeAt(i)>127) {
return false;
   }
}

if (user.match(userPat)==null) {
return false;
}

var IPArray=domain.match(ipDomainPat);
if (IPArray!=null) {

for (var i=1;i<=4;i++) {
if (IPArray[i]>255) {
return false;
   }
}
return true;
}

var atomPat=new RegExp("^" + atom + "$");
var domArr=domain.split(".");
var len=domArr.length;
for (i=0;i<len;i++) {
if (domArr[i].search(atomPat)==-1) {
alert(errEMailValDomain);
return false;
   }
}


if (len<2) {
alert(errEMailValURL);
return false;
}

return true;
}

function formatNumber(num,decimalNum,bolLeadingZero,bolParens,bolCommas)
/**********************************************************************
	IN:
		NUM - the number to format
		decimalNum - the number of decimal places to format the number to
		bolLeadingZero - true / false - display a leading zero for
										numbers between -1 and 1
		bolParens - true / false - use parenthesis around negative numbers
		bolCommas - put commas as number separators.
 
	RETVAL:
		The formatted number!
 **********************************************************************/
{ 
	//Add first zero if value is only decimal
	if (num.indexOf('.') == 0) num = '0' + num;
	
	if (isNaN(parseInt(num))) return "NaN";

	var tmpNum = num;
	var iSign = num < 0 ? -1 : 1;		// Get sign of number
	
	// Adjust number so only the specified number of numbers after
	// the decimal point are shown.
	tmpNum *= Math.pow(10,decimalNum);
	tmpNum = Math.round(Math.abs(tmpNum))
	tmpNum /= Math.pow(10,decimalNum);
	tmpNum *= iSign;					// Readjust for sign
	
	
	// Create a string object to do our formatting on
	var tmpNumStr = new String(tmpNum);

	// See if we need to strip out the leading zero or not.
	
	if (!bolLeadingZero && num < 1 && num > -1 && num != 0)
		if (num > 0)
			tmpNumStr = tmpNumStr.substring(1,tmpNumStr.length);
		else
			tmpNumStr = "-" + tmpNumStr.substring(2,tmpNumStr.length);
		
	// See if we need to put in the commas
	if (bolCommas && (num >= 1000 || num <= -1000)) {
		var iStart = tmpNumStr.indexOf(".");
		if (iStart < 0)
			iStart = tmpNumStr.length;

		iStart -= 3;
		while (iStart >= 1) {
			tmpNumStr = tmpNumStr.substring(0,iStart) + "," + tmpNumStr.substring(iStart,tmpNumStr.length)
			iStart -= 3;
		}		
	}
		
	//Complete format zeroes
	if (tmpNumStr.indexOf('.') == -1 && decimalNum > 0) tmpNumStr += '.';
	var addZero = decimalNum-(tmpNumStr.length-tmpNumStr.indexOf('.'));
	for (var i = 0;i<=addZero;i++)
	{
		tmpNumStr += '0';
	}

	// See if we need to use parenthesis
	if (bolParens && num < 0)
		tmpNumStr = "(" + tmpNumStr.substring(1,tmpNumStr.length) + ")";
	
	//Add first zero if value is only decimal
	if (tmpNumStr.indexOf('.') == 0) tmpNumStr = '0' + tmpNumStr;

	return tmpNumStr;		// Return our formatted string!
}

function Mid(str, start, len)
{
        if (start < 0 || len < 0) return "";

        var iEnd, iLen = String(str).length;
        if (start + len > iLen)
                iEnd = iLen;
        else
                iEnd = start + len;

        return String(str).substring(start,iEnd);
}

function Left(str, n){
	if (n <= 0)
	    return "";
	else if (n > String(str).length)
	    return str;
	else
	    return String(str).substring(0,n);
}

function Right(str, n){
    if (n <= 0)
       return "";
    else if (n > String(str).length)
       return str;
    else {
       var iLen = String(str).length;
       return String(str).substring(iLen, iLen - n);
    }
}

function chkMax(e, f, m)
{
	if(f.value.length == m && (e.keyCode != 8 && e.keyCode != 9 && e.keyCode != 35 && e.keyCode != 36 && e.keyCode != 37 
	&& e.keyCode != 38 && e.keyCode != 39 && e.keyCode != 40 && e.keyCode != 46 && e.keyCode != 16))return false; else return true;
}

function IsNumeric(sText)
{
   var ValidChars = "0123456789.-";
   var IsNumber=true;
   var Char;

 
   for (i = 0; i < sText.length && IsNumber == true; i++) 
      { 
      Char = sText.charAt(i); 
      if (ValidChars.indexOf(Char) == -1) 
         {
         IsNumber = false;
         }
      }
   return IsNumber;
   
}


function GetXmlHttpObject(handler)
{ 
	var objXmlHttp=null;

	if (navigator.userAgent.indexOf("MSIE")>=0)
	{ 
		var strName="Msxml2.XMLHTTP";
		if (navigator.appVersion.indexOf("MSIE 5.5")>=0)
		{
			strName="Microsoft.XMLHTTP";
		} 
		try
		{ 
			objXmlHttp=new ActiveXObject(strName);
			objXmlHttp.onreadystatechange=handler;
			return objXmlHttp;
		} 
		catch(e)
		{ 
			alert("Error. Scripting for ActiveX might be disabled");
			return;
		} 
	} 
	if (navigator.userAgent.indexOf("Mozilla")>=0 || navigator.userAgent.indexOf("Opera")>=0)
	{
		objXmlHttp=new XMLHttpRequest();
		objXmlHttp.onload=handler;
		objXmlHttp.onerror=handler;
		return objXmlHttp;
	}
} 

function setSelectionRange(input, selectionStart, selectionEnd) {
  if (input.setSelectionRange) {
    input.focus();
    input.setSelectionRange(selectionStart, selectionEnd);
  }
  else if (input.createTextRange) {
    var range = input.createTextRange();
    range.collapse(true);
    range.moveEnd('character', selectionEnd);
    range.moveStart('character', selectionStart);
    range.select();
  }
}

function replaceSelection (input, replaceString) {
	if (input.setSelectionRange) {
		var selectionStart = input.selectionStart;
		var selectionEnd = input.selectionEnd;
		input.value = input.value.substring(0, selectionStart)+ replaceString + input.value.substring(selectionEnd);
    
		if (selectionStart != selectionEnd){ 
			setSelectionRange(input, selectionStart, selectionStart + 	replaceString.length);
		}else{
			setSelectionRange(input, selectionStart + replaceString.length, selectionStart + replaceString.length);
		}

	}else if (document.selection) {
		var range = document.selection.createRange();

		if (range.parentElement() == input) {
			var isCollapsed = range.text == '';
			range.text = replaceString;

			 if (!isCollapsed)  {
				range.moveStart('character', -replaceString.length);
				range.select();
			}
		}
	}
}


// We are going to catch the TAB key so that we can use it, Hooray!
function catchTab(item,e){
	if(navigator.userAgent.match("Gecko")){
		c=e.which;
	}else{
		c=e.keyCode;
	}
	if(c==9){
		replaceSelection(item,String.fromCharCode(9));
		setTimeout("document.getElementById('"+item.id+"').focus();",0);	
		return false;
	}
		    
}