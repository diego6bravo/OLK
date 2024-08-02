function valFrm()
{
	var frm = document.frmOps;
	if (frm.opName.value == '')
	{
		alert(valOpName);
		frm.opName.focus();
		return false;
	}
	if (frm.Filter.value != '' && frm.valFilter.value == 'Y')
	{
		alert(valOpFilter);
		frm.Filter.focus();
		return false;
	}
	return true;
}
function valFrmDet()
{
	var frm = document.frmOpsDet;
	if (frm.LineID)
	{
		if (frm.LineID.length)
		{
			for (var i = 0;i<frm.LineID.length;i++)
			{

				var aliasDesc = document.getElementById('aliasDesc' + LineID[i].value);
				if (aliasDesc.value == '')
				{
					alert(valFldDesc);
					aliasDesc.focus();
					return false;
				}
			}
		}
		else
		{
			var aliasDesc = document.getElementById('aliasDesc' + LineID.value);
			if (aliasDesc.value == '')
			{
				alert(valFldDesc);
				aliasDesc.focus();
				return false;
			}
		}
	}
	return true;
}
function valFrmFld()
{
	var frm = document.frmOpsFld;
	if (frm.fldID.value == '')
	{
		alert(valFld);
		frm.fldID.focus();
		return false;
	}
	if (frm.fldID.value == 'Custom' || frm.LineID.value != '')
	{
		if (frm.AliasDesc.value == '')
		{
			alert(valFldDesc);
			frm.AliasDesc.focus();
			return false;
		}
		if (frm.AliasID.value == '' || frm.valFld.value == 'Y')
		{
			alert(valFldQry);
			frm.AliasDesc.focus();
			return false;
		}
	}	
	return true;
}

function doExpand(id)
{
	var sign = document.getElementById('signExpand' + id);
	
	var tr = document.getElementById('tr' + id);
	var display = false;

	if (tr.style.display == 'none')
	{
		tr.style.display = '';
		display = true;
	}
	else
	{
		tr.style.display = 'none';
		display = false;
	}
		
	sign.innerHTML = (display ? '[-]' : '[+]');
}


function changeFld(fld)
{
	document.getElementById('trFldQry').style.display = fld == 'Custom' ? '' : 'none';
	document.getElementById('AliasDesc').disabled = fld != 'Custom';
	document.getElementById('AliasDesc').style.borderColor = fld == 'Custom' ? '' : '#848284';
	
	var styleID = document.getElementById('StyleID');
	
	styleID.disabled = fld == 'Custom';
	document.getElementById('tdApply').style.display = fld == 'Custom' ? '' : 'none';
	
	if (fld != 'Custom' && parseInt(document.frmOps.Operation.value) != 5)
	{
		var arr = fld.split('{S}');
		var enableCheck = arr[1] == 'Y';
		var typeID = arr[2];
		
		for (var i = styleID.options.length - 1;i>1;i--)
		{
			styleID.options.remove(i);
		}
		
		if (enableCheck)
		{
			styleID.options[styleID.options.length] = new Option(txtCheckBox, 2);
		}
		
		if (typeID == 'A')
		{
			styleID.options[styleID.options.length] = new Option(txtVendorSelector, 3);
		}
	}
}

function changeOp(op)
{
	document.getElementById('trTargetObj').style.display = op == 2 || op == 3 || op == 6 ? '' : 'none';
	document.getElementById('trGenNewDoc').style.display = op == 6 ? '' : 'none';
}

var verfyButton;
var hdverfyButton;
function VerfyFilter()
{
	verfyButton = document.frmOps.btnVerfyFilter;
	hdverfyButton = document.frmOps.valFilter;

	document.frmVerfyQuery.type.value = 'OpFilter';
	document.frmVerfyQuery.Query.value = document.frmOps.Filter.value;
	document.frmVerfyQuery.Operation.value = document.frmOps.Operation.value;
	document.frmVerfyQuery.ObjectID.value = document.frmOps.ObjectID.value;
	document.frmVerfyQuery.submit();
}

function VerfyFld()
{
	verfyButton = document.frmOpsFld.btnVerfyFld;
	hdverfyButton = document.frmOpsFld.valFld;

	document.frmVerfyQuery.type.value = 'OpFld';
	document.frmVerfyQuery.Query.value = document.frmOpsFld.AliasID.value;
	document.frmVerfyQuery.Operation.value = document.frmOps.Operation.value;
	document.frmVerfyQuery.ObjectID.value = document.frmOps.ObjectID.value;
	document.frmVerfyQuery.submit();
}

function VerfyQueryVerified()
{
	verfyButton.src='images/btnValidateDis.gif';
	verfyButton.style.cursor = '';
	hdverfyButton.value='N';
}
