

function createPayment()
{
	setFlowAlertVars('R1', '', 'javascript:window.location.href=\'payments/newDocGo.asp?AddPath=\';', '');
	doFlowAlert();
}
function createDocument(objCode)
{
	if (parseInt(objCode) != 48)
		setFlowAlertVars('D1', objCode, 'javascript:window.location.href=\'ventas/newDocGo.asp?obj=' + objCode + '&AddPath=\';', '');
	else
		setFlowAlertVars('D1', '48', 'javascript:window.location.href=\'ventas/newCashInv.asp?AddPath=\';', '');
	doFlowAlert();
}
/* Advanced Custom Search */
function goAdSearch(ID, ObjID)
{
	 doMyLink('adCustomSearch.asp', 'ID=' + ID + '&adObjID=' + ObjID, '');
}

function showAdSearch(tbl, pageID)
{
	clearTimeout(curAdSearchTimerID);
	
	myTbl = document.getElementById('tblAdSearch');

	//myTbl.style.left = GetLeftPos(tbl) + (rtl == '' ? 96 : -1857);
	myTbl.style.left = rtl == '' ? 106 : -71;
	
	myTbl.style.top = GetTopPos(tbl)-153 - f_scrollTop();
	
	myTbl.style.display = '';
}
var ignoreClearAdSearch = false;
var curAdSearchTimerID; 

function hideAdSearch()
{
	curAdSearchTimerID = setTimeout("clearAdSearch()", 500);
}

function clearAdSearch()
{
	if (!ignoreClearAdSearch)
	{
		document.getElementById('tblAdSearch').style.display = 'none';
	}
	else
	{
		ignoreClearAdSearch = false;
	}
}

/* End */ 
var ns = (navigator.appName.indexOf("Netscape") != -1);
var d = document;
var px = document.layers ? "" : "px";
var isShowRep = false;
function JSFX_FloatDiv(id, sx, sy)
{
	var el=d.getElementById?d.getElementById(id):d.all?d.all[id]:d.layers[id];
	window[id + "_obj"] = el;
	if(d.layers)el.style=el;
	el.cx = el.sx = sx;el.cy = el.sy = sy;
	el.sP=function(x,y){this.style.left=x+px;this.style.top=y+px;};
	el.flt=function()
	{
		var addX = 0;
		if (rtl != '')
		{
			addX = ns ? innerWidth : 
			document.documentElement && document.documentElement.clientWidth ? 
			document.documentElement.clientWidth : document.body.clientWidth;
		}

		var pX, pY;
		pX = (this.sx >= 0) ? 0 : ns ? innerWidth : 
		document.documentElement && document.documentElement.clientWidth ? 
		document.documentElement.clientWidth : document.body.clientWidth;
		pY = ns ? pageYOffset : document.documentElement && document.documentElement.scrollTop ? 
		document.documentElement.scrollTop : document.body.scrollTop;
		if(this.sy<0) 
		pY += ns ? innerHeight : document.documentElement && document.documentElement.clientHeight ? 
		document.documentElement.clientHeight : document.body.clientHeight;
		this.cx += (pX + this.sx - this.cx)/8;this.cy += (pY + this.sy - this.cy)/8;

		this.sP(this.sx + addX, this.cy);
		setTimeout(this.id + "_obj.flt()", 40);
	}
	return el;
}

function setMinRepSize()
{
	document.getElementById('trMinRep').style.display = '';
	var h = document.getElementById('tblMinRep').offsetHeight;
	if (parseInt(h) < 200)
	{
		if (document.getElementById('minRepMoveUp')) document.getElementById('minRepMoveUp').style.display = 'none';
		if (document.getElementById('minRepMoveDown')) document.getElementById('minRepMoveDown').style.display = 'none';
		if (document.getElementById('scrollMinRep1')) document.getElementById('scrollMinRep1').height = h;
		if (document.getElementById('scrollMinRep2')) document.getElementById('scrollMinRep2').height = h;
		if (document.getElementById('scrollMinRep3')) document.getElementById('scrollMinRep3').style.height = h;
	}
	
}

var minRepScrollVal = 0;
var minRepScrollStop = false;
function minRepMove(dir)
{
	switch (dir)
	{
		case 'U':
			minRepScrollVal = 1;
			break;
		case 'D':
			minRepScrollVal = -1;
			break;
	}
	minRepScrollStop = true;
	setTimeout('scrollMinRep();', 1);
}

var minRepTimer = -1;
function loadMinRep()
{
	if (minRepTimer != -1) clearTimeout(minRepTimer);
	if (document.getElementById('tblMinRepWait'))
	{
		document.getElementById('tblMinRepWait').style.display = '';
		document.getElementById('imgMinRepWait').style.display = '';
		minRepTimer = setTimeout('doLoadMinRep();', 250);
	}
}

function doLoadMinRep()
{
	var curDate = new Date();
	var url='ventas/getMinRepData.asp?date=' + curDate.getTime();

	xmlHttp=GetXmlHttpObject(setMinRepData);
	xmlHttp.open("GET", url , true);
	xmlHttp.send(null);
}

function setMinRepData()
{
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
	{
		var strRetVal = xmlHttp.responseText;
		if (strRetVal != '')
		{
			var arrVals = strRetVal.split('{S}');
			for (var i = 0;i<arrVals.length;i++)
			{
				var ID = arrVals[i].split('{=}')[0];
				var Value = arrVals[i].split('{=}')[1];
				document.getElementById('mrVal' + ID).innerHTML = Value + '&nbsp;';
			}
		}
		document.getElementById('tblMinRepWait').style.display = 'none';
		document.getElementById('imgMinRepWait').style.display = 'none';
		minRepTimer = -1;
	}
}

function scrollMinRep()
{
	if (minRepScrollStop)
	{
		var tbl = $('#tblMinRep');
		
		var top = parseInt(tbl.css('top').replace('px', ''))+minRepScrollVal;

		if (minRepScrollVal == 1 && top > 0) top = 0;
		else if (minRepScrollVal == -1 && ((top*-1)+200) > tbl.height) top = (tbl.height-200)*-1;
		
		tbl.css('top', top + 'px'); 
		
		setTimeout('scrollMinRep();', 1);
	}
}

function stopMinRepMove()
{
	minRepScrollStop = false;
}

var NewValueTradField;
function doFldTrad(Table, ColumnID, ID, ColumnName, Type, NewValue)
{
	if (NewValue != null) NewValueTradField = NewValue;
	page = '';
	document.frmAdminTrad.Table.value = Table;
	document.frmAdminTrad.ColumnID.value = ColumnID;
	document.frmAdminTrad.ID.value = ID;
	document.frmAdminTrad.ColumnName.value = ColumnName;
	document.frmAdminTrad.Type.value = Type;
	if (NewValue != null) document.frmAdminTrad.NewValue.value = NewValue.value;
	else document.frmAdminTrad.NewValue.value = '';
	document.frmAdminTrad.IsNew.value = (NewValue == null ? 'N' : 'Y');
	switch (Type)
	{
		case 'T':
			w = 400;
			h = 82;
			break;
		case 'M':
			w = 400;
			h = 144;
			break;
		case 'R':
			w = 640;
			h = 480;
			break;
	}
	OpenWin = this.open(page, "CtrlWindow", "toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=no, width=" + w + ",height=" + h);
	document.frmAdminTrad.submit();
}
function setNewFldTrad(NewValue)
{
	NewValueTradField.value = NewValue;
	NewValueTradField = null;
}

function doItemFilter(item)
{
	document.frmSmallSearch.string.value = item;
	goSearchCart();
	document.frmSmallSearch.submit();
}

function showHideSection(tdIcon, trData)
{
	if (trData.style.display == 'none')
	{
		trData.style.display = '';
		tdIcon.innerText = '[-]';
	}
	else
	{
		trData.style.display = 'none';
		tdIcon.innerText = '[+]';
	}
}

function showUDF()
{
	trUDF.style.display = '';
	tdShowUDF.innerText = '[-]';
}

if (typeof doNoLang == 'undefined')
{
	doSelLang();
	document.onclick=clearSelectLang;
}


function chkMax(e, f, m)
{
	if(f.value.length == m && (e.keyCode != 8 && e.keyCode != 9 && e.keyCode != 35 && e.keyCode != 36 && e.keyCode != 37 
	&& e.keyCode != 38 && e.keyCode != 39 && e.keyCode != 40 && e.keyCode != 46 && e.keyCode != 16))
	{
		window.status = txtValFldMaxChar.replace('{0}', m);
		setTimeout("clear()",3000) 
		return false;
	}
	else { return true; }
}

function clear () { 
window.status= ""; 
} 

function doBlink() 
	{
	var blink = document.all.tags("BLINK")
	for (var i=0; i < blink.length; i++)
		blink[i].style.visibility = blink[i].style.visibility == "" ? "hidden" : ""
    }

function startBlink() 
	{
    if (document.all)
		setInterval("doBlink()",2500)
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

function chStyle(Cell, Type) { 
if (Type == 1) { var varColor = "#0075ea" } else { var varColor = "" }
Cell.style.backgroundColor = varColor
}

disableChkWin = false;
function chkWin() 
{ 
	try
	{
		if (OpenWin != null) if (!disableChkWin) if (!OpenWin.closed) OpenWin.focus() 
	}
	catch(err) {}
}

function clearWin() { OpenWin = null; }

function printStory(divId, SelDes) {

	var trCXCCmpName = document.getElementById('trCXCCmpName') && document.getElementById('hdTitle');
	
	w=window.open('','newwin')
	w.document.write('<html ' + curDir + '><head><link rel="stylesheet" type="text/css" href="design/' + SelDes + '/style/stylenuevo.css"></head><body onLoad="window.print()">');
   
	if (trCXCCmpName)
	{
		w.document.write('<table cellpadding="0" cellspacing="0" border="0" width="100%">' + document.getElementById('hdTitle').value + '</table>');
	}
	
	w.document.write(document.getElementById(divId).innerHTML);
   
	w.document.write('</body></html>');

	controls = w.document.getElementsByTagName('input');
	for (var i = 0;i<controls.length;i++)
	{
		if (controls[i].attributes.item('type').value == 'button' || controls[i].attributes.item('type').value == 'submit')
		{
			controls[i].style.display = 'none';
		}
	}
	
	if (trCXCCmpName)
	{
		w.document.getElementById('trCXCCmpName').style.display = 'none';
	}
	
	w.document.close();
}

var winl = (screen.width - 300) / 2;
var wint = (screen.height - 200) / 2;

function saveInPwd(status) { window.open('cartPdf.asp?status=' + status); }


function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layerh&&i<d.layerh.length;i++) x=MM_findObj(n,d.layerh[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}
function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function goRep(rsIndex, varsCount)
{
	document.frmGoRep.rsIndex.value = rsIndex;
	if (varsCount == 0)
	{
		document.frmGoRep.target = '';
		document.frmGoRep.action = 'report.asp';
		document.frmGoRep.AddPath.value = '';
		document.frmGoRep.pop.value = '';
		document.frmGoRep.submit();
	}
	else
	{
		document.frmGoRep.cmd.value = '';
		document.frmGoRep.target = 'RepVals';
		document.frmGoRep.action = 'portal/viewRepVals.asp';
		document.frmGoRep.AddPath.value = '';
		document.frmGoRep.pop.value = 'Y';
		OpenWin = this.open('', 'RepVals', "scrollbars=Yes,resizable=no, width=368,height=402");
		OpenWin.focus();
		document.frmGoRep.submit();
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
alert(txtValEMailDomain);
return false;
   }
}


if (len<2) {
alert(txtEMailValidURL);
return false;
}

return true;
}

/******* Admin Trad Form Java Script Start *******/

var NewValueTradField;
function doFldTrad(Table, ColumnID, ID, ColumnName, Type, NewValue)
{
	if (NewValue != null) NewValueTradField = NewValue;
	page = '';
	document.frmAdminTrad.Table.value = Table;
	document.frmAdminTrad.ColumnID.value = ColumnID;
	document.frmAdminTrad.ID.value = ID;
	document.frmAdminTrad.ColumnName.value = ColumnName;
	document.frmAdminTrad.Type.value = Type;
	if (NewValue != null) document.frmAdminTrad.NewValue.value = NewValue.value;
	else document.frmAdminTrad.NewValue.value = '';
	document.frmAdminTrad.IsNew.value = (NewValue == null ? 'N' : 'Y');
	switch (Type)
	{
		case 'T':
			w = 400;
			h = 82;
			break;
		case 'M':
			w = 400;
			h = 144;
			break;
		case 'R':
			w = 640;
			h = 480;
			break;
	}
	OpenWin = this.open(page, "CtrlWindow", "toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=no, width=" + w + ",height=" + h);
	document.frmAdminTrad.submit();
}
function setNewFldTrad(NewValue)
{
	NewValueTradField.value = NewValue;
	NewValueTradField = null;
}
/******* Admin Trad Form Java Script End *******/


/******* Small Export Java Script Start *******/
var expTblID;
function showSmallExpTbl()
{
	var tbl = document.getElementById('tblSmallExport');
	var img = document.getElementById('imgVentasExport');
	
	if (rtl == '')
	{
		tbl.style.left = GetLeftPos(img)-4 + 'px';
	}
	else
	{
		tbl.style.left = -109;
	}
	tbl.style.top = GetTopPos(img)-127 + 'px';
	tbl.style.display = '';
}
function clearSmallExpTbl()
{
	var tbl = document.getElementById('tblSmallExport');
	tbl.style.display = 'none';
	MM_swapImage('iconos_ventas_r2_c2','','ventas/images/iconos_ventas_export.gif',1);
	expTblID=window.clearInterval(expTblID);
}
function setSmallExpTDBorder(td, over)
{
	if (over)
	{	
		td.style.borderTop='1px #005CBA solid';
		td.style.borderLeft='1px #005CBA solid';
		td.style.borderRight='1px #1975cf solid';
		td.style.borderBottom='1px #1975cf solid';
	}
	else
	{
		td.style.borderTop='';
		td.style.borderLeft='';
		td.style.borderRight='';
		td.style.borderBottom='';
	}
}
/******* Small Expot Java Script End *******/


var ns6=document.getElementById&&!document.all
var ie=document.all

function changeto(e,highlightcolor){
source=ie? event.srcElement : e.target
if (source.tagName=="TR"||source.tagName=="TABLE")
return
while(source.tagName!="TD"&&source.tagName!="HTML")
source=ns6? source.parentNode : source.parentElement
if (source.style.backgroundColor!=highlightcolor&&source.id!="ignore")
source.style.backgroundColor=highlightcolor
}

function contains_ns6(master, slave) { //check if slave is contained by master
while (slave.parentNode)
if ((slave = slave.parentNode) == master)
return true;
return false;
}

function changeback(e,originalcolor){
if
(ie&&(event.fromElement.contains(event.toElement)||source.contains(event.toElement)||source.id=="ignore")||source.tagName=="TR"||source.tagName=="TABLE")
return
else if (ns6&&(contains_ns6(source, e.relatedTarget)||source.id=="ignore"))
return
if (ie&&event.toElement!=source||ns6&&e.relatedTarget!=source)
source.style.backgroundColor=originalcolor
}
