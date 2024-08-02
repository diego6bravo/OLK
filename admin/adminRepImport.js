
function valFrmUpload(frm)
{
	var myFile = frm.xmlFile.value;
	if (myFile.substring(myFile.length-3).toLowerCase() != 'xml')
	{
		alert(txtValXmlFile);
		return false;
	}
	ignoreUnload = true;
	return true;
}

function valFrmImp()
{
	if (document.frmImport.rsIndex.length)
	{
		var chk = false;
		for (var i = 0;i<document.frmImport.rsIndex.length;i++)
		{
			if (document.frmImport.rsIndex[i].checked)
			{
				chk = true;
				if (!ValidateRS(document.frmImport.rsIndex[i].value))
				{
					return false;
				}
			}
		}
		if (!chk)
		{
			alert(txtValSelRep);
			return false;
		}
	}
	else
	{
		if (!document.frmImport.rsIndex.checked)
		{
			alert(txtValSelRep);
			return false;
		}
		else if (!ValidateRS(document.frmImport.rsIndex.value))
		{
			return false;
		}
	}
	ignoreUnload = true;
	return true;
}

function ValidateRS(rsIndex)
{
	if (document.getElementById('rsName' + rsIndex).value == '')
	{
		alert(txtValRepName);
		document.getElementById('rsName' + rsIndex).focus();
		return false;
	}
	else if (document.getElementById('rgIndex' + rsIndex).selectedIndex == 0)
	{
		alert(txtValRepGrp);
		document.getElementById('rgIndex' + rsIndex).focus();
		return false;
	}
	return true;
}

function chkAll(col, chk)
{
	if (document.getElementById(col).length)
		for (var i = 0;i<document.getElementById(col).length;i++)
		{
			if (!document.getElementById(col)[i].disabled)
			document.getElementById(col)[i].checked = chk;
		}
	else
	{
		if (!document.getElementById(col).disabled)
		document.getElementById(col).checked = chk;
	}
}

function chkCheckAll(col)
{
	var chk = true;
	if (document.getElementById(col).length)
	{
		for (var i = 0;i<document.getElementById(col).length;i++)
		{
			if (!document.getElementById(col)[i].checked && !document.getElementById(col)[i].disabled)
			{
				chk = false;
				break;
			}
		}
	}
	else
	{
		if (!document.getElementById(col).disabled)
		chk = document.getElementById(col).checked;
	}
		
	document.getElementById('chkAll' + col).checked = chk;
}
