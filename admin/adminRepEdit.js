doSelFontLayer();
doSelFontSizeLayer();
document.onclick=clearSelectFont;

function clearColor(fld, prv)
{
	fld.value = '';
	prv.bgColor='';
}

function blinkIt() {
 if (!document.all) return;
 else {
   for(i=0;i<document.all.tags('blink').length;i++){
      s=document.all.tags('blink')[i];
      s.style.visibility=(s.style.visibility=='visible')?'hidden':'visible';
   }
 }
}

function addAlterColor(rsIndex, ColorID, LineNum)
{
	window.location.href='adminRepEdit.asp?rsIndex=' + rsIndex + '&ColorID=' + ColorID + '&AlterOf=Y&LineNum=' + LineNum + '&repCmd=repColor&#editColor';
}

function doEditColor(rsIndex, ColorID, LineID, LineNum, AlterNum)
{
	window.location.href='adminRepEdit.asp?rsIndex=' + rsIndex + '&repCmd=repColor&ColorID=' + ColorID + '&LineID=' + LineID + '&LineNum=' + LineNum + '&AlterNum=' + AlterNum + '&#editColor';
}

function moveColor(dir, rsIndex, ColorID, LineID)
{
	window.location.href='repSubmit.asp?rsIndex=' + rsIndex + '&ColorID=' + ColorID + '&LineID=' + LineID + '&Cmd=moveColor&dir=' + dir;
}

function doDelColor(rsIndex, ColorID, LineID)
{
	if (LineID == 0)
	{
		var msgStr = txtValDelColor1.replace('{0}', (ColorID+1)) + '<br>' + txtValDelColor2;
		var msgClickYes = 'window.location.href=\'repSubmit.asp?cmd=remRSCol&rsIndex=' + rsIndex + '&ColorID=' + ColorID + '\';';
		var msgClickNo = 'window.location.href=\'repSubmit.asp?cmd=remRSCol&rsIndex=' + rsIndex + '&ColorID=' + ColorID + '&LineID=' + LineID + '\';';
							
		showMsgBox(msgStr, 2, msgClickYes, msgClickNo, '', 1);
	}
	else
	{
		var msgStr = txtConfDelColor.replace('{0}', (ColorID+1)).replace('{1}', (LineID+1));
		var msgClickYes = 'window.location.href=\'repSubmit.asp?cmd=remRSCol&rsIndex=' + rsIndex + '&ColorID=' + ColorID + '&LineID=' + LineID + '\';';
		
		showMsgBox(msgStr, 1, msgClickYes, '', '', 1);
	}
}


function ChangeOp(op)
{
	if (op == 'N' || op == 'NN')
	{
		if (!document.frmAddEditCol.colOpBy.disabled)
		{
			document.frmAddEditCol.colOpBy.disabled = true;
			document.frmAddEditCol.colValCol.disabled = true;
			document.frmAddEditCol.colValVal.disabled = true;
			document.frmAddEditCol.colValDat.disabled = true;
		}
	}
	else
	{
		if (document.frmAddEditCol.colOpBy.disabled)
		{
			document.frmAddEditCol.colOpBy.disabled = false;
			document.frmAddEditCol.colValCol.disabled = false;
			document.frmAddEditCol.colValVal.disabled = false;
			document.frmAddEditCol.colValDat.disabled = false;
		}
	}
}

function ChangeColName(val)
{
	if (document.frmAddEditCol.colOpBy.value == 'V')
	{
		if (val.split('{|}')[1] != 'D' && document.getElementById('tblValDat').style.display!='none')
		{
			document.getElementById('tblValDat').style.display = 'none';
			document.frmAddEditCol.colValVal.style.display = '';
		}
		else if (val.split('{|}')[1] == 'D')
		{
			document.getElementById('tblValDat').style.display = '';
			document.frmAddEditCol.colValVal.style.display = 'none';
		}
	}
}

function ChangeOpBy(val)
{
	if (val == 'F')
	{
		txtOpByVal.style.display='none';
		txtOpByFld.style.display='';
		if (document.frmAddEditCol.colName.value.split('{|}')[1] != 'D') document.frmAddEditCol.colValVal.style.display='none';
		else document.getElementById('tblValDat').style.display='none';
		document.frmAddEditCol.colValCol.style.display='';
	}
	else if (val == 'V')
	{
		txtOpByVal.style.display='';
		txtOpByFld.style.display='none';
		if (document.frmAddEditCol.colName.value.split('{|}')[1] != 'D') document.frmAddEditCol.colValVal.style.display='';
		else document.getElementById('tblValDat').style.display='';
		document.frmAddEditCol.colValCol.style.display='none';

	}
}

function valValVal(fld)
{
	if (document.frmAddEditCol.colName.value.split('{|}')[1] == 'N' && fld.value != '')
	{
		if (!IsNumeric(fld.value))
		{
			alert(clearHTMLChar(txtValNumVal));
			fld.value = '';
			fld.focus();
			return false;
		}
		return true;
	}
	return true;
}


function valFrmAdmCol()
{
	arrColID = document.frmAdmColors.ColID;
	if (arrColID)
	{
		if (arrColID.length)
		{
			for (var i = 0;i<arrColID.length;i++)
			{
				ColID = arrColID[i].value
				if (document.getElementById('ColAlias' + ColID).value == '')
				{
					alert(txtValAlias);
					document.getElementById('ColAlias' + ColID).focus();
					return false;
				}
			}
		}
		else
		{
			ColID = arrColID.value
			if (document.getElementById('ColAlias' + ColID).value == '')
			{
				alert(txtValAlias);
				document.getElementById('ColAlias' + ColID).focus();
				return false;
			}
		}
	}
	return true;
}

function valFrmAddEditCol()
{
	if (document.frmAddEditCol.Alias.value == '')
	{
		alert(clearHTMLChar(txtValAlias));
		document.frmAddEditCol.Alias.focus();
		return false;
	}
	else if (document.frmAddEditCol.Active.checked)
	{
		if (document.frmAddEditCol.colOp.value != 'N' && document.frmAddEditCol.colOp.value != 'NN')
		{
 			if (document.frmAddEditCol.colOpBy.value == 'V')
 			{
	 			if (document.frmAddEditCol.colName.value.split('{|}')[1] != 'D')
 				{
	 				if (document.frmAddEditCol.colValVal.value == '')
	 				{
		 				alert(clearHTMLChar(txtValOpValue));
		 				document.frmAddEditCol.colValVal.focus();
		 				return false;
	 				}
	 				else if (!valValVal(document.frmAddEditCol.colValVal))
	 				{
	 					return false;
	 				}
 				}
	 			else if (document.frmAddEditCol.colValDat.value == '')
	 			{
	 				alert(clearHTMLChar(txtValOpValue));
	 				document.getElementById('btnValDatImg').click();
	 				return false;
	 			}
 			}
 			else if (document.frmAddEditCol.colOpBy.value == 'F')
 			{
 				if (document.frmAddEditCol.colValCol.selectedIndex == 0)
 				{
	 				alert(clearHTMLChar(txtValFldOp));
	 				document.frmAddEditCol.colValCol.focus();
	 				return false;
	 			}
	 			else if (document.frmAddEditCol.colName.value.split('{|}')[1] != document.frmAddEditCol.colValCol.value.split('{|}')[1])
	 			{
	 				alert(clearHTMLChar(txtValFldTypes).replace('{0}', document.frmAddEditCol.colValCol.value.split('{|}')[0]).replace('{1}', document.frmAddEditCol.colName.value.split('{|}')[0]));
	 				document.frmAddEditCol.colValCol.selectedIndex = 0;
	 				return false;
	 			}
 			}
 		}
	}
	return true;
}
