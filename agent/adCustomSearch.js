function chkNumValue(e)
{
	switch (e.keyCode)
	{
		case 37: //Left
		case 39: //Right
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

function getValue(myType, fld) {
if (fld.value == '') { return; } 
	updFld = fld;
	if (fld.value.indexOf('*') == -1) {
		document.frmGetValue.Type.value = myType;
		document.frmGetValue.searchStr.value = fld.value;
		document.frmGetValue.submit();
	}
	else { launchSelect(myType, fld.value); }
}
function launchSelect(myType, Value){
	var retVal = window.showModalDialog('topGetValueSelect.asp?Type=' + myType + '&Value=' + Value,'','dialogWidth:500px;dialogHeight:500px');
	if (retVal != '' && retVal != null){
		updFld.value = retVal; setTargetVal(retVal); retVal = '';
	} 
	else { 
		updFld.value = '';
	}
}
function setValue(src, value, myType){
	if (value != '') 
	{ updFld.value = value; setTargetVal(value); }
	else { if(src == 0)launchSelect(myType, updFld.value); }
}
function setTargetVal(value)
{
	if (Right(updFld.name, 4) == "From")
	{
		setFldName = Left(updFld.name, (updFld.name.length-4));
		fldTo = document.getElementById(setFldName + 'To');
		if (fldTo.value == '') { fldTo.value = value; fldTo.select(); }
	}
}
var noVal = false;
function reload(TargetID)
{
	noVal = true;
	if (TargetID != '')
	{
		var arrIndex = TargetID.toString().split(', ');
		for (var i = 0;i<arrIndex.length;i++)
		{
			document.getElementById('var' + arrIndex[i]).value = '';
		}
	}
	document.frmVars.isSubmit.value = "R";
	document.frmVars.action = 'adCustomSearch.asp';
	document.frmVars.submit();
}
function chkNum(fld, dType)
{
	if (dType != 'nvarchar')
	{
		if (!myIsNumeric(fld.value))
		{
			alert(txtValNumVal);
			fld.value = '';
			fld.focus();
		}
		else if (dType == 'int')
		{
			fld.value = parseInt(fld.value);
		}
	}
}


var OpenWin = this;
var Field;
function chkWin() { if (!OpenWin.closed) OpenWin.focus() }

function Start(o, page, w, h, s, r) {
Field = o
OpenWin = this.open(page, "queryWin", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable="+r+", width="+w+",height="+h);
OpenWin.focus()
}

function setTimeStamp(Nothing, var1) {
	Field.value = var1;
	if (Field.onchange != null) Field.onchange();
}
