function myTrim(sString) 
{
	var retVal = sString;
	
	try
	{
		retVal = retVal.replace(/^\s+|\s+$/g, '');
	}
	catch(ex)
	{
		retVal = retVal.replace('&nbsp;', '').replace(' ', '');
	}
	
	return retVal;
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


function MyIsNumeric(sText)
{
   var ValidChars = "0123456789.";
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


function formatNumberDec(num, places, comma){ 
var isNeg=0; 
if(num < 0) { 
num=num*-1; 
isNeg=1; 
} 
var myDecFact = 1; 
var myPlaces = 0; 
var myZeros = ""; 
while(myPlaces < places) { 
myDecFact = myDecFact * 10; 
myPlaces = eval(myPlaces) + eval(1); 
myZeros = myZeros + "0"; 
} 
onum=Math.round(num*myDecFact)/myDecFact; 
integer=Math.floor(onum); 
if (Math.ceil(onum) == integer) { 
decimal=myZeros; 
} else{ 
decimal=Math.round((onum-integer)* myDecFact) 
} 
decimal=decimal.toString(); 
if (decimal.length<places) { 
fillZeroes = places - decimal.length; 
for (z=0;z<fillZeroes;z++) { 
decimal="0"+decimal; 
} 
} 

if(places > 0) { 
decimal = "." + decimal; 
} 
if(comma == 1) { 
integer=integer.toString(); 
var tmpnum=""; 
var tmpinteger=""; 
var y=0; 
for (x=integer.length;x>0;x--) { 
tmpnum=tmpnum+integer.charAt(x-1); 
y=y+1; 
if (y==3 & x>1) { 
tmpnum=tmpnum+","; 
y=0; 
} 
} 

for (x=tmpnum.length;x>0;x--) { 
tmpinteger=tmpinteger+tmpnum.charAt(x-1); 
} 


finNum=tmpinteger+""+decimal; 
} else { 
finNum=integer+""+decimal; 
} 

if(isNeg == 1) { 
finNum = "-" + finNum; 
} 

return finNum; 
} 
