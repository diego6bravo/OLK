var tblAcct;
var tblAcctID;

function clearAccount(id)
{
	document.getElementById(id).value = '';
	document.getElementById(id + 'Disp').value = '';
	document.getElementById(id + 'Name').value = '';
}

function clearSelectAccount() { if (tblAcct) { tblAcct.style.display='none'; clearSelectLang(); } }

document.onclick=clearSelectAccount;

function showSelectAccount(img, id, t, e)
{
	if (tblAcct) tblAcct.style.display = 'none';
	
	var tblID = 'lySelAct' + t;

	tblAcct = document.all ? document.getElementById(tblID) : document.getElementById(tblID);
	
	if (!tblAcct)
	{
	
		$.post('accountControlFetch.asp', { Type: t }, function(data) { doSelActLayer(img, id, t, e, data); });
		return true;
	}
	
	if (tblAcct.style.display == 'none')
	{
		var offset = $('#imgCmb' + id).offset();
		tblAcct.style.left = offset.left - 407;
		tblAcct.style.top = offset.top + 15;
		tblAcct.style.display = '';
		e.cancelBubble = true;
		tblAcctID = id;
	}
	
	return true;
}

function doSelActLayer(img, id, t, e, data)
{
	var clearSpace = document.all ? document.getElementById('clearSpace') : document.getElementById('clearSpace');
	
	var strTbl = "<div style=\"border-left:1px solid #68A6C0; border-right:1px solid #68A6C0; border-bottom:1px solid #68A6C0; position: absolute; width: 420px; height: 200px; overflow: scroll; overflow-x: hidden; z-index: 1; background-color:#D9F0FD; display: none; left:10px; top:431px\" id=\"lySelAct" + t + "\"> " + 
	"<table border=\"0\" cellspacing=\"0\" width=\"100%\"> " + 
	"	<tr> " + 
	"		<td onmouseover=\"bgColor='#EBF8FE'\" onmouseout=\"bgColor=''\" style=\"cursor: default; font-face: Verdana; font-size: 10px; color: #3F7B96; \" onclick=\"setSelAcct('', '', '');\">&nbsp;</td> " + 
	"	</tr> ";
	
	var arrData = data.split('{S}');
	for (var i = 0;i<arrData.length;i++)
	{
		var arrCol = arrData[i].split('{C}');
		
		strTbl += "<tr><td onmouseover=\"bgColor='#EBF8FE'\" onmouseout=\"bgColor=''\" style=\"cursor: default; font-face: Verdana; font-size: 10px; color: #3F7B96;\" onclick=\"setSelAcct('" + arrCol[0] + "', '" + arrCol[1].replace("'", "\'") + "', '" + arrCol[2] + "');\"><font face=\"Arial\">" + arrCol[2] + " - " + arrCol[1] + "</font></td></tr>";
		
	} 
	
	strTbl += "</table></div>";
	
	clearSpace.innerHTML += strTbl;
	
	showSelectAccount(img, id, t, e);
}

function setSelAcct(acctCode, acctName, acctDisp)
{
	var txtSel = document.all ? document.getElementById('txtSel' + tblAcctID) : document.getElementById('txtSel' + tblAcctID);
	$('#' + tblAcctID).val(acctCode);
	txtSel.innerText = acctDisp + ' - ' + acctName;
}