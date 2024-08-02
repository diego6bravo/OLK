function Start2(theURL, popW, popH, type) { // V 1.0
var winleft = (screen.width - popW) / 2;
var winUp = (screen.height - popH) / 2;
winProp = 'width='+popW+',height='+popH+',left='+winleft+',top='+winUp+',toolbar=no,scrollbars=yes,menubar=no,location=no,resizable=no'
theURL2 = theURL+'?update='+type+'&pop=Y&rParent=Y'
OpenWin = window.open(theURL2, "CtrlWindow2", winProp)
}
var OpenWin = this;
function chkWin() { if (OpenWin) { OpenWin.focus(); } }
function clearWin() { OpenWin = null; } 

function clearCur(value)
{
	return parseFloat(getNumeric(value.replace(DocCur,'')));
}

function SetTotal()
{
	document.Form1.Total.value = DocCur + ' ' + OLKFormatNumber((clearCur(document.Form1.Total.value) + clearCur(document.Form1.SaldoPag.value)),SumDec);
	updatePayment();
}
function SetSaldo()
{
	var openBal = clearCur(document.Form1.ImpInc.value)-clearCur(document.Form1.Pagado.value);
	document.Form1.SaldoPag.value = DocCur + ' ' + OLKFormatNumber(openBal,SumDec);
	document.Form1.pagVal.value = clearCur(document.Form1.Pagado.value);
}
function updatePayment()
{
	if (clearCur(document.Form1.Total.value) < 0) document.Form1.Total.value = DocCur + ' ' + OLKFormatNumber(0,SumDec);
	document.Form1.Pagado.value = DocCur + ' ' + OLKFormatNumber(getNumeric(vPagado)-getNumeric(cashSum)+clearCur(document.Form1.Total.value),SumDec);
	var openBal = clearCur(document.Form1.ImpInc.value)-clearCur(document.Form1.Pagado.value);
	document.Form1.SaldoPag.value = DocCur + ' ' + OLKFormatNumber(openBal,SumDec);
	document.Form1.pagVal.value = clearCur(document.Form1.Pagado.value);
}

function chkNum(Field)
{
	if (!MyIsNumeric(getNumeric(Field.value)))
	{
		alert(txtValNumVal);
		Field.value = DocCur + ' ' + OLKFormatNumber(0,SumDec);
	}
	else
	{
		if (parseFloat(getNumeric(Field.value)) < 0)
		{
			alert(txtValNumMinVal.replace("{0}", "0"));
			Field.value = DocCur + ' ' + OLKFormatNumber(0,SumDec);
		}
		else
		{
			Field.value = DocCur + ' ' + OLKFormatNumber(getNumeric(Field.value),SumDec);
		}
	}
}
