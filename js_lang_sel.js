function showSelectLang(btn, e)
{
	if (myLng.indexOf(', ') == -1) 
	{
		e.cancelBubble = true;
		return false;
	}
	
	if (tblSelLng.style.display == 'none')
	{
		tblSelLng.style.left = GetLeftPos(btn)- ((rtl != 'rtl/' && !jsLangRev || rtl == 'rtl/' && jsLangRev) ? 164 : 0);
		tblSelLng.style.top = GetTopPos(btn)+15;
		tblSelLng.style.display = '';
		e.cancelBubble = true;
		return true;
	}
	else
	{
		tblSelLng.style.display = 'none';
		return false;
	}
}

function doChangeLang(lng)
{
	document.frmChangeLng.newLng.value = lng;
	document.frmChangeLng.submit();
}

function doSelLang()
{
	var arr = myLng.split(', ');
	if (arr.length > 1)
	{
		var addStyle = '';
		var strTbl = 
		"<table border=\"0\" id=\"tblSelLng\" cellspacing=\"0\" width=\"180\" id=\"table1\" style=\"font-family: Verdana; font-size: 10px; position: absolute; z-index: 1; display: none; left:10px; top:431px\"> "
		
		for (var i = 0;i<arr.length;i++)
		{
			if (i+1 == arr.length)
				addStyle = ';border-bottom:1px solid ' + jsLangCol1;
				
			var tdSmallLng =
			"		<td align=\"center\" width=\"20\" style=\"border-left: 1px solid " + jsLangCol1 + "; border-right-width: 1px; border-top: 1px solid " + jsLangCol1 + "" + addStyle + "\"> " + 
			"		<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" bgcolor=\"" + jsLangCol1 + "\" style=\"cursor: hand\" onclick=\"showSelectLang(this, event);\"> " + 
			"			<tr> " + 
			"				<td width=\"16\" height=\"16\" align=\"center\"> " + 
			"				<font size=\"1\" face=\"Verdana\" color=\"" + jsLangCol4 + "\">" + arr[i].split('{S}')[0].toUpperCase() + "</font></td> " + 
			"			</tr> " + 
			"		</table> " + 
			"		</td> ";
			
			var tdLng = 
			"		<td style=\"border-left-width: 1px; border-right: 1px solid " + jsLangCol1 + "; border-top: 1px solid " + jsLangCol1 + "" + addStyle + "\"> " + 
			"		<font color=\"" + jsLangCol5 + "\">" + arr[i].split('{S}')[1] + "</font></td> ";
			
			strTbl += 
			"	<tr bgcolor=\"" + jsLangCol2 + "\" style=\"cursor: hand\" onmouseout=\"this.bgColor='" + jsLangCol2 + "'\" onmouseover=\"this.bgColor='" + jsLangCol3 + "'\" onclick=\"doChangeLang(\'" + arr[i].split('{S}')[0] + "\')\"> " + 
			(rtl != 'rtl/' ? tdSmallLng + tdLng : tdLng + tdSmallLng) +
			"	</tr> "
		}
		strTbl += "</table> ";
		document.write(strTbl);
	}
}

function clearSelectLang()
{
	if (myLng.indexOf(', ') == -1) return;
	tblSelLng.style.display = 'none';
}