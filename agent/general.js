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

function getNumeric(value)
{
	var retVal = value;
	retVal = retVal.replace(GetFormatSep, '');
	retVal = retVal.replace(GetFormatComma, '.');
	return retVal;
}

function getNumericVB(value)
{
	var retVal = value;
	retVal = retVal.replace(GetFormatSep, '');
	return retVal;
}

function valKeyNum(e)
{
	switch (e.keyCode)
	{
		case 8:
		case 9:
		case 16:
		case 36:
		case 46:
		case 48:
		case 49:
		case 50:
		case 51:
		case 52:
		case 53:
		case 54:
		case 55:
		case 56:
		case 57:
		case 96:
		case 97:
		case 98:
		case 99:
		case 100:
		case 101:
		case 102:
		case 103:
		case 104:
		case 105:
		case 37: // Left
		case 39: //Right
			return true;
	}
	return false;
}

function valKeyNumDec(e)
{
	switch (e.keyCode)
	{
		case 8:
		case 9:
		case 16:
		case 38:
		case 36:
		case 46:
		case 48:
		case 49:
		case 50:
		case 51:
		case 52:
		case 53:
		case 54:
		case 55:
		case 56:
		case 57:
		case 96:
		case 97:
		case 98:
		case 99:
		case 100:
		case 101:
		case 102:
		case 103:
		case 104:
		case 105:
		case 37: // Left
		case 39: //Right
		case 188: //comma
		case 190: //dot
		case 110: //dot
			return true;
	}
	return false;
}

String.prototype.trim = function() {
	return this.replace(/^\s+|\s+$/g,"");
}
String.prototype.ltrim = function() {
	return this.replace(/^\s+/,"");
}
String.prototype.rtrim = function() {
	return this.replace(/\s+$/,"");
}

function Pic(page, w, h, s, r) {
	OpenWin = this.open('', 'ImageThumb', 'scrollbars='+s+',resizable='+r+', width='+w+',height='+h);
	doMyLink(page.split('?')[0], page.split('?')[1], 'ImageThumb');
	OpenWin.focus();
}


function OLKFormatNumber(value, decimals)
{
	var snum = new String(value);
	
	var sec = snum.split('.');
	
	var whole = parseFloat(sec[0]);
	
	var result = '';
	
	if(sec.length > 1){
		var dec = new String(sec[1]);
		dec = String(parseFloat(sec[1])/Math.pow(10,(dec.length - decimals)));
		dec = String(whole + Math.round(parseFloat(dec))/Math.pow(10,decimals));
		var dot = dec.indexOf('.');
		if(dot == -1){
			dec += '.'; 
			dot = dec.indexOf('.');
		}
		while(dec.length <= dot + decimals) { dec += '0'; }
		result = dec;
	} 
	else
	{
		var dot;
		var dec = new String(whole);
		dec += '.';
		dot = dec.indexOf('.');        
		while(dec.length <= dot + decimals) { dec += '0'; }
		result = dec;
	}
	
	if (GetFormatComma != '.') result = result.replace('.', GetFormatComma);
	
	if (decimals == 0)
	{
		result = result.substring(0, result.length-1);
	}
	
	result += '';
	x = result.split('.');
	x1 = x[0];
	x2 = x.length > 1 ? '.' + x[1] : '';
	var rgx = /(\d+)(\d{3})/;
	while (rgx.test(x1)) {
		x1 = x1.replace(rgx, '$1' + GetFormatSep + '$2');
  	}
	return x1 + x2;
}

function selViewType(value, viewCount)
{
	document.getElementById('varViewTypeID' + viewCount).value = value;
	
	checkViewType(document.getElementById('tdViewTypeT' + viewCount), (value == 'T'), viewCount);
	checkViewType(document.getElementById('tdViewTypeC' + viewCount), (value == 'C'), viewCount);
	checkViewType(document.getElementById('tdViewTypeL' + viewCount), (value == 'L'), viewCount);
}
function setDateFormat(fld)
{
	if (strDateSep != '/') fld.value = fld.value.replace('/', strDateSep).replace('/', strDateSep);
}
function checkViewType(td, enable, viewCount)
{
	td.style.borderBottomColor = enable ? document.getElementById('varAlterColor' + viewCount).value : 'transparent';
}

function chkNext(o, e, b)
{
	switch (e.keyCode)
	{
		case 38:
			if (parseInt(b) > 1) o[parseInt(b)-2].focus();
			return false;
			break;
		case 40:
			if (o.length)
			{
				if (o[parseInt(b)]) o[parseInt(b)].focus();
			}
			return false;
			break;
		case 8:
		case 9:
		case 16:
		case 38:
		case 36:
		case 46:
		case 48:
		case 49:
		case 50:
		case 51:
		case 52:
		case 53:
		case 54:
		case 55:
		case 56:
		case 57:
		case 96:
		case 97:
		case 98:
		case 99:
		case 100:
		case 101:
		case 102:
		case 103:
		case 104:
		case 105:
		case 190: //Punto decimal
		case 110: //Punto decimal
		case 188: //Comma decimal
		case 37: // Left
		case 39: //Right
			return true;
	}
	return false;
}


function chkNextCat(o, s, e, b, cols)
{
	switch (e.keyCode)
	{
		case 37: //Left
			if (parseInt(b) > 1)
				if (getCursorPos(s) == 0)
					o[parseInt(b)-2].focus();
			return true;
		case 39: //Right
			if (o.length)
			{
				if (getCursorPos(s) == s.value.length)
					if (o[parseInt(b)]) o[parseInt(b)].focus();
			}
			return true;
		case 38:
			if (parseInt(b-cols) > 1) 
				o[parseInt(b-cols)-2].focus();
			return false;
		case 40:
			if (o.length)
				if (o[parseInt(b+cols)]) o[parseInt(b+cols)].focus();
			break;
		case 8:
		case 9:
		case 16:
		case 38:
		case 36:
		case 46:
		case 48:
		case 49:
		case 50:
		case 51:
		case 52:
		case 53:
		case 54:
		case 55:
		case 56:
		case 57:
		case 96:
		case 97:
		case 98:
		case 99:
		case 100:
		case 101:
		case 102:
		case 103:
		case 104:
		case 105:
		case 190: //Punto decimal
		case 110: //Punto decimal
		case 188: //Comma decimal
			return true;
	}
	return false;
}


function getCursorPos(o) 
{
	var curVal = o.value;

	var objRange = document.selection.createRange();
	var curRange = objRange.text;

	var custString = '#%~';

	objRange.text = curRange + custString; objRange.moveStart('character', (0 - curRange.length - custString.length));

	var newVal = o.value;

	objRange.text = curRange;

	for (i=0; i <= newVal.length; i++) 
	{
		var sTemp = newVal.substring(i, i + custString.length);
		if (sTemp == custString) 
		{
			var cursorPos = (i - curRange.length);
			return cursorPos;
		}
	}
}


function f_clientWidth() {
	return f_filterResults (
		window.innerWidth ? window.innerWidth : 0,
		document.documentElement ? document.documentElement.clientWidth : 0,
		document.body ? document.body.clientWidth : 0
	);
}
function f_clientHeight() {
	return f_filterResults (
		window.innerHeight ? window.innerHeight : 0,
		document.documentElement ? document.documentElement.clientHeight : 0,
		document.body ? document.body.clientHeight : 0
	);
}
function f_scrollLeft() {
	return f_filterResults (
		window.pageXOffset ? window.pageXOffset : 0,
		document.documentElement ? document.documentElement.scrollLeft : 0,
		document.body ? document.body.scrollLeft : 0
	);
}
function f_scrollTop() {
	return f_filterResults (
		window.pageYOffset ? window.pageYOffset : 0,
		document.documentElement ? document.documentElement.scrollTop : 0,
		document.body ? document.body.scrollTop : 0
	);
}
function f_filterResults(n_win, n_docel, n_body) {
	var n_result = n_win ? n_win : 0;
	if (n_docel && (!n_result || (n_result > n_docel)))
		n_result = n_docel;
	return n_body && (!n_result || (n_result > n_body)) ? n_body : n_result;
}

function isChkboxChecked(fld)
{
	var retVal = false;
	if (fld)
	{
		if (fld.length)
		{
			for (var i = 0;i<fld.length;i++)
			{
				if (fld[i].checked)
				{
					retVal = true;
					break;
				}
			}
		}
		else
		{
			retVal = fld.checked;
		}
	}
	return retVal;
}

function myTrim(value)
{
	if (value != '') return value.replace(/^\s+|\s+$/g, '');
	else return '';
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


function goWL(doc, DefCatOrdr, MainDoc)
{
	var OrderBy = DefCatOrdr == 'C' ? 'OITM.ItemCode' : 'ItemName';
	var frmGoWL = 	'<form method="post" action="search.asp" name="frmGoWL">' +
					'<input type="hidden" name="cmd" value="wish">' +
					'<input type="hidden" name="document" value="C">' +
					'<input type="hidden" name="orden1" value="' + OrderBy + '">' +
					'<input type="hidden" name="orden2" value="asc">' +
					'<input type="hidden" name="chkWL" value="Y">' +
					'</form>';
	doc.write(frmGoWL);
	doc.frmGoWL.submit();
}

function goProm(doc, DefCatOrdr, MainDoc)
{
	var OrderBy = DefCatOrdr == 'C' ? 'OITM.ItemCode' : 'ItemName';
	var frmGoProm = 	'<form method="post" action="prom.asp" name="frmGoProm">' +
					'<input type="hidden" name="document" value="C">' +
					'<input type="hidden" name="orden1" value="' + OrderBy + '">' +
					'<input type="hidden" name="orden2" value="asc">' +
					'<input type="hidden" name="chkProm" value="Y">' +
					'</form>';
	doc.write(frmGoProm);
	doc.frmGoProm.submit();
}

function setAllSize()
{
	if (document.getElementById('tdNavI'))
	{
		curH = 0;
		if (tdNavI.length)
		{
			for (var i = 0;i<tdNavI.length;i++)
			{
				if (tdNavI[i].clientHeight > curH)
					curH = tdNavI[i].clientHeight;
			}
			for (var i = 0;i<tdNavI.length;i++)
			{
				tdNavI[i].style.height = curH + 'px';
			}
		}
	}
	
	if (document.getElementById('tdSecI'))
	{
		curH = 0;
		if (tdSecI.length)
		{
			for (var i = 0;i<tdSecI.length;i++)
			{
				if (tdSecI[i].clientHeight > curH)
					curH = tdSecI[i].clientHeight;
			}
			for (var i = 0;i<tdSecI.length;i++)
			{
				tdSecI[i].style.height = curH + 'px';
			}
		}
	}
	
	if (document.getElementById('tdCIName'))
	{
		curH = 0;
		if (tdCIName.length)
		{
			for (var i = 0;i<tdCIName.length;i++)
			{
				if (tdCIName[i].clientHeight > curH)
					curH = tdCIName[i].clientHeight;
			}
			for (var i = 0;i<tdCIName.length;i++)
			{
				tdCIName[i].style.height = curH + 'px';
			}
		}
	}
}

function GetTopPos(inputObj)
{
	
  var returnValue = inputObj.offsetTop;
  while((inputObj = inputObj.offsetParent) != null){
  	returnValue += inputObj.offsetTop;
  }
  return returnValue;
}

function GetLeftPos(inputObj)
{
  var returnValue = inputObj.offsetLeft;
  while((inputObj = inputObj.offsetParent) != null)returnValue += inputObj.offsetLeft;
  return returnValue;
}

function MyIsNumeric(sText)
{
   var ValidChars = "0123456789.,";
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
function dateAddExtention(p_Interval, p_Number){ 
 
 
   var thing = new String(); 
    
    
   //in the spirt of VB we'll make this function non-case sensitive 
   //and convert the charcters for the coder. 
   p_Interval = p_Interval.toLowerCase(); 
    
   if(isNaN(p_Number)){ 
    
      //Only accpets numbers  
      //throws an error so that the coder can see why he effed up    
      throw "The second parameter must be a number. \n You passed: " + p_Number; 
      return false; 
   } 
 
   p_Number = new Number(p_Number); 
   switch(p_Interval.toLowerCase()){ 
      case "yyyy": {// year 
         this.setFullYear(this.getFullYear() + p_Number); 
         break; 
      } 
      case "q": {      // quarter 
         this.setMonth(this.getMonth() + (p_Number*3)); 
         break; 
      } 
      case "m": {      // month 
         this.setMonth(this.getMonth() + p_Number); 
         break; 
      } 
      case "y":      // day of year 
      case "d":      // day 
      case "w": {      // weekday 
         this.setDate(this.getDate() + p_Number); 
         break; 
      } 
      case "ww": {   // week of year 
         this.setDate(this.getDate() + (p_Number*7)); 
         break; 
      } 
      case "h": {      // hour 
         this.setHours(this.getHours() + p_Number); 
         break; 
      } 
      case "n": {      // minute 
         this.setMinutes(this.getMinutes() + p_Number); 
         break; 
      } 
      case "s": {      // second 
         this.setSeconds(this.getSeconds() + p_Number); 
         break; 
      } 
      case "ms": {      // second 
         this.setMilliseconds(this.getMilliseconds() + p_Number); 
         break; 
      } 
      default: { 
       
         //throws an error so that the coder can see why he effed up and 
         //a list of elegible letters. 
         throw   "The first parameter must be a string from this list: \n" + 
               "yyyy, q, m, y, d, w, ww, h, n, s, or ms.  You passed: " + p_Interval; 
         return false; 
      } 
   } 
   return this; 
} 
Date.prototype.dateAdd = dateAddExtention; 

function olkOpenObj(obj, entry, id2)
{
	doMyLink('cxcDocDetailOpen.asp', 'DocEntry=' + entry + '&high=' + id2 + '&DocType=' + obj + '&pop=Y', '_blank');
}

