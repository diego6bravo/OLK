var curMnu = null;
var curMnuTD = null;
var curSubMnu = null;
var curSubMnuTD = null;
var curMnuTimerID = null;
var curSubMnuTimerID = null;

function showMnu(td, mnu, isSub)
{
	var getFldPos;
	if (!isSub)
	{
		timeHideMnu();
		getFldPos = td;
	}
	else
	{
		clearTimeout(curMnuTimerID);
		timeHideSubMnu(false);
		getFldPos = document.getElementById('imgSubMnu' + mnu);
	}
	
	td.bgColor = '#0075EA';
	myGenMen = document.getElementById('myGenMen' + mnu);
	myGenMen.style.display = '';
	if (myGenMen.offsetWidth < td.offsetWidth) myGenMen.width = td.offsetWidth;
	myGenMen.style.left = GetLeftPos(getFldPos)+ (!isSub ? 0 : 10)+(rtl == '' ? 0 : getFldPos.offsetWidth-myGenMen.offsetWidth)-(isSub && rtl != '' ? 21 : 0);
	myGenMen.style.top = GetTopPos(getFldPos)+ (!isSub ? 20 : -7);
	if (!isSub)
	{
		curMnu = mnu;
		curMnuTD = td;
	}
	else
	{
		curSubMnu = mnu;
		curSubMnuTD = td;
		myGenMen.style.zIndex = 2;
	}
}

function hideMnu(td)
{
	curMnuTimerID = setTimeout("timeHideMnu()", 500);
}

function showMnuItm(td, isSub)
{
	if (curSubMnu != null && !isSub)
	{
		timeHideSubMnu(false);
	}
	else if (isSub)
	{
		clearTimeout(curSubMnuTimerID);
	}
	td.bgColor = '#0075EA';
	clearTimeout(curMnuTimerID);
}

function hideMnuItm(td)
{
	td.bgColor = '#0066CB';
	curMnuTimerID = setTimeout("timeHideMnu()", 500);
}

function hideSubMnu()
{
	curSubMnuTimerID = setTimeout("timeHideSubMnu(true)", 500);
}

function timeHideSubMnu(clearMnu)
{
	if (curSubMnuTimerID != null) clearTimeout(curSubMnuTimerID);
	if (curSubMnu != null)
	{
		curSubMnuTD.bgColor = '#0066CB';
		document.getElementById('myGenMen' + curSubMnu).style.display = 'none';
		curSubMnu = null;
		curSubMnuTD = null;
		if (clearMnu) timeHideMnu();
	}
}

function timeHideMnu()
{
	if (curMnuTimerID != null) clearTimeout(curMnuTimerID);
	if (curMnu != null)
	{
		timeHideSubMnu(false);
		curMnuTD.bgColor = '';
		document.getElementById('myGenMen' + curMnu).style.display = 'none';
		curMnu = null;
		curMnuTD = null;
		
	}
}

function loadMenus()
{
	for (var i = 0;i<menus.length;i++)
	{
		var mnu = menus[i].items;
		
		var myMenu = "<table border=\"0\" cellspacing=\"0\" cellpadding=\"0\" bgcolor=\"#00488F\" id=\"myGenMen" + menus[i].label + "\" style=\"position: absolute; display: none; z-index: 1\"> " + 
		"<tr> " + 
		"	<td> " + 
		"	<table border=\"0\" width=\"100%\" cellspacing=\"1\" cellpadding=\"0\"> ";
		
		for (var j = 0;j<mnu.length;j++)
		{
			var myMenuMouseOver = '';
			var myMenuMouseOut = '';
			var myMenuStyle = '';
			if (!mnu[j].hasSub)
			{
				myMenuMouseOver = 'showMnuItm(this, ' + menus[i].isSub + ')';
				myMenuMouseOut = 'hideMnuItm(this)';
				myMenuStyle = '';	
				myMenuOnClick = mnu[j].action;
			}
			else
			{
				myMenuMouseOver = 'showMnu(this, \'' + mnu[j].subMenu + '\', true)';
				myMenuMouseOut = 'hideSubMnu()';
				myMenuStyle = 'cursor: default';				
				myMenuOnClick = '';
			}
			myMenu += 
			"		<tr> " + 
			"			<td bgcolor=\"#0066CB\" onmouseover=\"" + myMenuMouseOver + ";\" onmouseout=\"" + myMenuMouseOut + ";\" onclick=\"" + myMenuOnClick + "\"> " + 
			"			<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" width=\"100%\" class=\"mySubMnuItem\" style=\"" + myMenuStyle + "\"> " + 
			"				<tr> " + 
			"					<td>" + mnu[j].label + "</td> ";
			
			if (mnu[j].hasSub)
			{
				myMenu += "<td valign=\"bottom\" align=\"" + (rtl == '' ? 'right' : 'left') + "\"><img id=\"imgSubMnu" + mnu[j].subMenu + "\" src=\"images/" + rtl + "arrows_white.gif\"></td>";
			}
			
			myMenu += "				</tr> " + 
			"			</table> " + 
			"			</td> " + 
			"		</tr> ";
		}
		
		myMenu += 
		"	</table> " + 
		"	</td> " + 
		"</tr> " + 
		"</table> ";
		document.write(myMenu);
	}
}

function goLink(lnk)
{
	window.location.href = lnk;
}

function goLinkPop(lnk)
{
	var load = window.open(lnk,'pop','scrollbars=yes,menubar=yes,resizable=yes,toolbar=yes,location=yes,status=yes');
}

function Menu(label, isSub) 
{
	this.type = "Menu";
	this.label = label;
	this.isSub = isSub;
	this.items = new Array();
	this.addMenuItem = addMenuItem;
	if (!window.menus) window.menus = new Array();
	window.menus[window.menus.length] = this;
}

function MenuItem(label, action, hasSub, subMenu)
{
	this.label = label;
	this.action = action;
	this.hasSub = hasSub;
	this.subMenu = subMenu;
}

function addMenuItem(newMenuItem) {
	this.items[this.items.length] = newMenuItem;
}
