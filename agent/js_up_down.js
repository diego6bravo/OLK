var NumUDField;
var NumUDTimerID = 0;
var NumUDMinVal = 0;
function doNumUD(udFld, dir)
{
	NumUDField = udFld;
	if (dir == 'U')
	{
		NumUDTimerID = setTimeout("doNumUDUp()", 250);
	}
	else
	{
		NumUDTimerID = setTimeout("doNumUDDown()", 250);
	}
}

function doNumUDUp()
{
	if(parseFloat(NumUDField.value)<32767)NumUDField.value=parseFloat(NumUDField.value)+1;
	NumUDTimerID = setTimeout("doNumUDUp()", 250);
}

function doNumUDDown()
{
	if(parseFloat(NumUDField.value)>NumUDMinVal)NumUDField.value=parseFloat(NumUDField.value)-1;
	NumUDTimerID = setTimeout("doNumUDDown()", 250);
}

function stopDoNumUD()
{
	clearTimeout(NumUDTimerID);
}

function doNumUDKeyDown(s, e)
{
	NumUDField = s;
	if (e.keyCode == 38)
	{
		doNumUDUp();
		stopDoNumUD();
		s.select();
		return false;
	}
	else if (e.keyCode == 40)
	{
		doNumUDDown();
		stopDoNumUD();
		s.select();
		return false;
	}
	else if (	e.keyCode >= 96 && e.keyCode <= 105 || e.keyCode >= 48 && e.keyCode <= 57 || 
				e.keyCode == 46 || e.keyCode == 36 || e.keyCode == 35 || e.keyCode == 8 || e.keyCode == 9 || e.keyCode == 16)
	{
		return true;
	}
	else
		return false;
}

function NumUDAttachMin(FormID, FieldID, btnUp, btnDown, MinVal)
{
	NumUDMinVal = MinVal;
	NumUDAttach(FormID, FieldID, btnUp, btnDown);
}

function NumUDAttach(FormID, FieldID, btnUp, btnDown)
{
	if (browserDetect() == 'firefox')
	{
		document.getElementById(btnUp).onmousedown = new Function('event','javascript:doNumUD(document.' + FormID + '.' + FieldID + ', \'U\');');
		document.getElementById(btnUp).onmouseup = new Function('event','javascript:stopDoNumUD();');
		document.getElementById(btnUp).onclick = new Function('event','javascript:if(parseFloat(document.' + FormID + '.' + FieldID + '.value)<32767)document.' + FormID + '.' + FieldID + '.value=parseFloat(document.' + FormID + '.' + FieldID + '.value)+1;');
		
		document.getElementById(btnDown).onmousedown = new Function('event','javascript:doNumUD(document.' + FormID + '.' + FieldID + ', \'D\');');
		document.getElementById(btnDown).onmouseup = new Function('event','javascript:stopDoNumUD();');
		document.getElementById(btnDown).onclick = new Function('event','javascript:if(parseFloat(document.' + FormID + '.' + FieldID + '.value)>0)document.' + FormID + '.' + FieldID + '.value=parseFloat(document.' + FormID + '.' + FieldID + '.value)-1;');
	
		document.getElementById(FieldID).onkeydown = new Function('event','javascript:return doNumUDKeyDown(this, event);');
	}
	else
	{
		document.getElementById(btnUp).onmousedown = new Function('javascript:doNumUD(document.' + FormID + '.' + FieldID + ', \'U\');');
		document.getElementById(btnUp).onmouseup = new Function('javascript:stopDoNumUD();');
		document.getElementById(btnUp).onclick = new Function('javascript:if(parseFloat(document.' + FormID + '.' + FieldID + '.value)<32767)document.' + FormID + '.' + FieldID + '.value=parseFloat(document.' + FormID + '.' + FieldID + '.value)+1;');
		
		document.getElementById(btnDown).onmousedown = new Function('javascript:doNumUD(document.' + FormID + '.' + FieldID + ', \'D\');');
		document.getElementById(btnDown).onmouseup = new Function('javascript:stopDoNumUD();');
		document.getElementById(btnDown).onclick = new Function('javascript:if(parseFloat(document.' + FormID + '.' + FieldID + '.value)>0)document.' + FormID + '.' + FieldID + '.value=parseFloat(document.' + FormID + '.' + FieldID + '.value)-1;');
	
		document.getElementById(FieldID).onkeydown = new Function('javascript:return doNumUDKeyDown(this, event);');
	}
	document.getElementById(FieldID).style.textAlign = 'right';
}