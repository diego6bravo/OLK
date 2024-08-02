var CartObject = -1;
function valFrmUpload(frm)
{
	var myFile = frm.xmlFile.value;
	if (myFile.substring(myFile.length-3).toLowerCase() != 'txt')
	{
		alert(txtValTxtFile);
		return false;
	}
	ignoreUnload = true;
	return true;
}

function ChkNum(Field, Min, Max, Val, setDec)
{
	if (!MyIsNumeric(getNumeric(Field.value)))
	{
		alert(txtValNumVal);
		Field.value = OLKFormatNumber(Val,setDec);
		return;
	}
	else if (parseFloat(getNumeric(Field.value)) <= parseFloat(getNumeric(Min)))
	{
		alert(txtValNumMinVal.replace("{0}", Min));
		Field.value = Man;
		return;
	}
	else if (Max != '')
	{
		if (parseFloat(getNumeric(Field.value)) > parseFloat(getNumeric(Max)))
		{
			alert(txtValNumMaxVal.replace("{0}", Max));
			Field.value = Max;
			return;
		}
	}
	
	Field.value = FormatNumber(getNumericVB(Field.value), setDec)
}
function setTblSet()
{
	tblImport.style.top = document.body.offsetHeight-25+document.body.scrollTop;
}


function setTotal(LineNum)
{
	var total = parseFloat(getNumeric(document.getElementById('Quantity' + LineNum).value))*
				parseFloat(getNumeric(document.getElementById('Price' + LineNum).value));
				
	if (UnEmbPriceSet)
	{
		var saleUnit = parseInt(document.getElementById('SaleUnit' + LineNum).value);
		if (saleUnit == 3) total = total * parseFloat(document.getElementById('SalPackUn' + LineNum).value);
	}
	document.getElementById('LineTotal' + LineNum).innerHTML = OLKFormatNumber(total, SumDec);
}
function chkAll(chkAll)
{
	if (document.frmImport.LineNum.length)
	{
		for (var i = 0;i<document.frmImport.LineNum.length;i++)
		{
			document.frmImport.LineNum[i].checked = chkAll;
		}
	}
	else
	{
		document.frmImport.LineNum.checked = chkAll;
	}
}
function chkCheckAll()
{
	if (document.frmImport.LineNum.length)
	{
		checked = true;
		for (var i = 0;i<document.frmImport.LineNum.length;i++)
		{
			if (!document.frmImport.LineNum[i].checked)
			{
				checked = false;
				break;
			}
		}
		document.frmImport.chkAllItems.checked = checked;
	}
	else
	{
		document.frmImport.chkAllItems.checked = document.frmImport.LineNum.checked;
	}
}
function valFrmImp()
{
	if (document.frmImport.LineNum.length)
	{
		var found = false;
		for (var i = 0;i<document.frmImport.LineNum.length;i++)
		{
			if (document.frmImport.LineNum[i].checked)
			{
				if (CartObject != 23)
				{
					var LineNum = document.frmImport.LineNum[i].value;
					var ItemCode = document.getElementById('ItemCode' + LineNum).value;
					if (parseFloat(getNumeric(document.getElementById('Quantity' + LineNum).value)) >
						parseFloat(getNumeric(document.getElementById('MaxQty' + LineNum).value)))
					{
						alert(txtValMaxQty.replace('{0}', ItemCode).replace('{1}', document.getElementById('MaxQty' + LineNum).value));
						return false;
					}
				}
				found = true;
			}
		}
		if (!found)
		{
			alert(txtValSelItms);
			return false;
		}
	}
	else
	{
		if (!document.frmImport.LineNum.checked)
		{
			alert(txtValSelItms);
			return false;
		}
		else
		{
			if (CartObject != 23)
			{
				var ItemCode = document.getElementById('ItemCode' + LineNum).value;
				if (parseFloat(getNumeric(document.getElementById('Quantity' + LineNum).value)) >
					parseFloat(getNumeric(document.getElementById('MaxQty' + LineNum).value)))
				{
					alert(txtValMaxQty.replace('{0}', ItemCode).replace('{1}', document.getElementById('MaxQty' + LineNum).value));
					return false;
				}
			}
		}
	}
	return true;
}
