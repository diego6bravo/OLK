function valFrm()
{
	if (clearCur(document.Form1.Total.value) != 0 && document.Form1.AcctCode.value == '')
	{
		alert(txtValSelAcct);
		return false;
	}
	
	return true;
}

function clearCur(value)
{
	return parseFloat(getNumeric(value.replace(DocCur,'')));
}

function SetSaldo()
{
	var openBal = clearCur(document.Form1.ImpInc.value)-clearCur(document.Form1.Pagado.value);
	document.Form1.SaldoPag.value = DocCur + ' ' + OLKFormatNumber(openBal,SumDec);
	document.Form1.pagVal.value = document.Form1.Pagado.value;
}
function updatePayment(Field)
{
	if (!MyIsNumeric(getNumeric(Field.value)))
	{
		alert(txtValNumVal)
		Field.value = 0;
		Field.focus();
	}
	else if (clearCur(document.Form1.Total.value) < 0)
	{
		alert(txtValNumMinVal.replace("{0}", "0"));
		Field.value = 0;
	}
	Field.value = DocCur + ' ' + OLKFormatNumber(getNumeric(Field.value),SumDec);
	
	document.Form1.Pagado.value = DocCur + ' ' + OLKFormatNumber(getNumeric(vPagado)-getNumeric(TrsfrSum)+clearCur(document.Form1.Total.value),SumDec);
	document.Form1.SaldoPag.value = DocCur + ' ' + OLKFormatNumber(clearCur(document.Form1.ImpInc.value)-clearCur(document.Form1.Pagado.value),SumDec);
	document.Form1.pagVal.value = document.Form1.Pagado.value
}
