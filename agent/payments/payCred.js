var InstalMent;
var var1;
var var2;
var var3;
var MinCredit;
var MinToPay;
var MaxValid;
var SPago = false

function clearWin()
{
	OpenWin = this;
}

var OpenWin = this;
function chkWin() { if (OpenWin) { OpenWin.focus(); } }

function Start(theURL, popW, popH, scroll) { // V 1.0
var winleft = (screen.width - popW) / 2;
var winUp = (screen.height - popH) / 2;
winProp = 'width='+popW+',height='+popH+',left='+winleft+',top='+winUp+',toolbar=no,scrollbars='+scroll+',menubar=no,location=no,resizable=no'
theURL2 = theURL+'?saldo='+document.Form2.SaldoPag.value+'&pop=Y&AddPath=../'
OpenWin = window.open(theURL2, "CtrlWindow2", winProp)
if (parseInt(navigator.appVersion) >= 4) { OpenWin.focus(); }

}
function textCounter(field, maxLimit) 
{
	if (field.value.length > maxLimit)
		field.value = field.value.substring(0, maxLimit);
}
function Ifthis(field,string) 
{
	if (field.value == string) 
	{
		field.value = '';
	}
}

function chkNum(Field, Max, Min, OldVal)
{
	var Value = Field.value;
	
	if (Value != "")
	{
		if (!MyIsNumeric(Value))
		{
			alert(txtValNumVal);
			if (OldVal != "") Field.value = OldVal;
			else Field.value = "";
			Field.focus();
			return;
		}
		
		if (parseFloat(getNumeric(Value)) < 0)
		{
			alert(txtValNumMinVal.replace("{0}", "0"));
			if (OldVal != "") Field.value = OldVal;
			else Field.value = "";
			Field.focus();
			return;
		}
		
		if (Max != null)
		{
			if (parseFloat(getNumeric(Value)) > Max)
			{
				alert(txtValNumMaxVal.replace("{0}", Max));
				Field.value = Max;
				Field.focus();
				return;
			}
		}
		
		if (Min != null)
		{
			if (parseFloat(getNumeric(Value)) < Min)
			{
				alert(txtValNumMinVal.replace("{0}", Min));
				Field.value = Min;
				Field.focus();
				return;
			}
		}
	}
}

function ClearCur(Value)
{
	if (Value != "")
	{
		return parseFloat(getNumeric(Value.replace(DocCur,"")));
	}
	else
		return 0;
}

function SetSaldo()
{
	var openBal = ClearCur(document.Form2.ImpInc.value)-ClearCur(document.Form2.Pagado.value);
	document.Form2.SaldoPag.value = DocCur + ' ' + OLKFormatNumber(openBal,SumDec);
	document.Form2.pagVal.value = document.Form2.Pagado.value;
}
function updatePayment()
{
	if (parseFloat(getNumeric(document.Form2.Total.value)) < 0) document.Form2.Total.value = 0;
	
	document.Form2.Pagado.value = DocCur + ' ' + OLKFormatNumber(getNumeric(Pagado)-getNumeric(creditsum)+getNumeric(document.Form2.Total.value),SumDec);
	document.Form2.SaldoPag.value = DocCur + ' ' + OLKFormatNumber(getNumeric(document.Form2.impinc.value)-getNumeric(document.form2.pagado.value),SumDec);
	document.Form2.pagVal.value = document.Form2.Pagado.value;
}
