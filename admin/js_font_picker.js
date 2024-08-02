
var fontUpdateField;
function setSelFont(Font)
{
	txtSelFontSample.innerHTML = '<font face=\"' + Font + '\">' + Font + '</font>';
	fontUpdateField.value = Font;
	lySelFont.style.display='none';
}

function setSelFontSize(Size)
{
	txtSelFontSizeSample.innerHTML = Size;
	fontUpdateField.value = Size;
	lySelFontSize.style.display='none';
}

function clearSelectFont() { lySelFont.style.display='none'; lySelFontSize.style.display='none'; clearSelectLang(); }

function showSelectFont(btn, updFld, e)
{
	if (lySelFont.style.display=='none')
	{
		clearSelectFont();
		fontUpdateField = updFld;
		lySelFont.style.left = GetLeftPos(btn)-107;
		lySelFont.style.top = GetTopPos(btn)+14;
		lySelFont.style.display='';
		e.cancelBubble = true;
		return true;
	}
	else
	{
		lySelFont.style.display='none';
		return false;
	}
}

function showSelectFontSize(btn, updFld, e)
{
	if (lySelFontSize.style.display=='none')
	{
		clearSelectFont();
		fontUpdateField = updFld;
		lySelFontSize.style.left = GetLeftPos(btn)-107;
		lySelFontSize.style.top = GetTopPos(btn)+14;
		lySelFontSize.style.display='';
		e.cancelBubble = true;
		return true;
	}
	else
	{
		lySelFontSize.style.display='none';
		return false;
	}
}
/*
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
}*/

function doSelFontLayer()
{
	document.write("<div style=\"border-left:1px solid #68A6C0; border-right:1px solid #68A6C0; border-bottom:1px solid #68A6C0; position: absolute; width: 120; z-index: 1;background-color:#D9F0FD; display: none; left:10px; top:431px\" id=\"lySelFont\"> " + 
	"<table border=\"0\" cellspacing=\"0\" width=\"100%\" id=\"table2\"> " + 
	"	<tr> " + 
	"		<td onmouseover=\"bgColor='#EBF8FE'\" onmouseout=\"bgColor=''\" style=\"cursor: default; font-size: 12px; color: #3F7B96; font-weight: bold\" onclick=\"setSelFont('&nbsp;');\">&nbsp;</td> " + 
	"	</tr> " + 
	"	<tr> " + 
	"		<td onmouseover=\"bgColor='#EBF8FE'\" onmouseout=\"bgColor=''\" style=\"cursor: default; font-size: 12px; color: #3F7B96; font-weight: bold\" onclick=\"setSelFont('Arial');\"><font face=\"Arial\">Arial</font></td> " + 
	"	</tr> " + 
	"	<tr> " + 
	"		<td onmouseover=\"bgColor='#EBF8FE'\" onmouseout=\"bgColor=''\" style=\"cursor: default; font-size: 12px; color: #3F7B96; font-weight: bold\" onclick=\"setSelFont('Times New Roman');\"><font face=\"Times New Roman\">Times New Roman</font></td> " + 
	"	</tr> " + 
	"	<tr> " + 
	"		<td onmouseover=\"bgColor='#EBF8FE'\" onmouseout=\"bgColor=''\" style=\"cursor: default; font-size: 12px; color: #3F7B96; font-weight: bold\" onclick=\"setSelFont('Courier');\"><font face=\"Courier\">Courier</font></td> " + 
	"	</tr> " + 
	"	<tr> " + 
	"		<td onmouseover=\"bgColor='#EBF8FE'\" onmouseout=\"bgColor=''\" style=\"cursor: default; font-size: 12px; color: #3F7B96; font-weight: bold\" onclick=\"setSelFont('Georgia');\"><font face=\"Georgia\">Georgia</font></td> " + 
	"	</tr> " + 
	"	<tr> " + 
	"		<td onmouseover=\"bgColor='#EBF8FE'\" onmouseout=\"bgColor=''\" style=\"cursor: default; font-size: 12px; color: #3F7B96; font-weight: bold\" onclick=\"setSelFont('Verdana');\"><font face=\"Verdana\">Verdana</font></td> " + 
	"	</tr> " + 
	"	<tr> " + 
	"		<td onmouseover=\"bgColor='#EBF8FE'\" onmouseout=\"bgColor=''\" style=\"cursor: default; font-size: 12px; color: #3F7B96; font-weight: bold\" onclick=\"setSelFont('Tahoma');\"><font face=\"Tahoma\">Tahoma</font></td> " + 
	"	</tr> " + 
	"</table> " + 
	"</div> ");
}


function doSelFontSizeLayer()
{
	document.write("<div style=\"border-left:1px solid #68A6C0; border-right:1px solid #68A6C0; border-bottom:1px solid #68A6C0; position: absolute; width: 120; z-index: 1;background-color:#D9F0FD; display: none; left:10px; top:431px\" id=\"lySelFontSize\"> " + 
	"<table border=\"0\" cellspacing=\"0\" width=\"100%\" id=\"table2\"> " + 
	/*"	<tr> " + 
	"		<td onmouseover=\"bgColor='#EBF8FE'\" onmouseout=\"bgColor=''\" style=\"cursor: default; font-size: 12px; color: #3F7B96\" onclick=\"setSelFontSize('&nbsp;');\">&nbsp;</td> " + 
	"	</tr> " + */
	"	<tr> " + 
	"		<td onmouseover=\"bgColor='#EBF8FE'\" onmouseout=\"bgColor=''\" style=\"cursor: default; font-size: 12px; color: #3F7B96\" onclick=\"setSelFontSize('1');\"><font face=\"Verdana\" size=\"1\">1</font></td> " + 
	"	</tr> " + 
	"	<tr> " + 
	"		<td onmouseover=\"bgColor='#EBF8FE'\" onmouseout=\"bgColor=''\" style=\"cursor: default; font-size: 12px; color: #3F7B96\" onclick=\"setSelFontSize('2');\"><font face=\"Verdana\" size=\"2\">2</font></td> " + 
	"	</tr> " + 
	"	<tr> " + 
	"		<td onmouseover=\"bgColor='#EBF8FE'\" onmouseout=\"bgColor=''\" style=\"cursor: default; font-size: 12px; color: #3F7B96\" onclick=\"setSelFontSize('3');\"><font face=\"Verdana\" size=\"3\">3</font></td> " + 
	"	</tr> " + 
	"	<tr> " + 
	"		<td onmouseover=\"bgColor='#EBF8FE'\" onmouseout=\"bgColor=''\" style=\"cursor: default; font-size: 12px; color: #3F7B96\" onclick=\"setSelFontSize('4');\"><font face=\"Verdana\" size=\"4\">4</font></td> " + 
	"	</tr> " + 
	"	<tr> " + 
	"		<td onmouseover=\"bgColor='#EBF8FE'\" onmouseout=\"bgColor=''\" style=\"cursor: default; font-size: 12px; color: #3F7B96\" onclick=\"setSelFontSize('5');\"><font face=\"Verdana\" size=\"5\">5</font></td> " + 
	"	</tr> " + 
	"	<tr> " + 
	"		<td onmouseover=\"bgColor='#EBF8FE'\" onmouseout=\"bgColor=''\" style=\"cursor: default; font-size: 12px; color: #3F7B96\" onclick=\"setSelFontSize('6');\"><font face=\"Verdana\" size=\"6\">6</font></td> " + 
	"	</tr> " + 
	"	<tr> " + 
	"		<td onmouseover=\"bgColor='#EBF8FE'\" onmouseout=\"bgColor=''\" style=\"cursor: default; font-size: 12px; color: #3F7B96\" onclick=\"setSelFontSize('7');\"><font face=\"Verdana\" size=\"7\">7</font></td> " + 
	"	</tr> " + 
	"</table> " + 
	"</div> ");
}
