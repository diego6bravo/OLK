var varLineID;

var retAcct;
function Start(page, retAction) {
retAcct = retAction
OpenWin = this.open(page, "DatePicker", "toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no, width=240,height=220");
}

function setTimeStamp(retAction, varDate) {
retAcct.value = varDate }

function setBranch(Branch, Account)
{
	if (varLineID == "new")
	{
		document.Form1.sucursal.value = Branch;
		document.Form1.cuenta.value = Account;
	}
	else
	{
		document.getElementById("sucursal" + varLineID).value = Branch;
		document.getElementById("cuenta" + varLineID).value = Account;
	}
}


var OpenWin = this;
function Start2(theURL, popW, popH, type) 
{
	var winleft = (screen.width - popW) / 2;
	var winUp = (screen.height - popH) / 2;	
	winProp = 'width='+popW+',height='+popH+',left='+winleft+',top='+winUp+',toolbar=no,scrollbars=yes,menubar=no,location=no,resizable=no'
	if (type != '') { theURL2 = theURL+'?update='+type+'&pop=Y&rParent=Y'; }
	else { theURL2 = theURL; }
	OpenWin = window.open(theURL2, "CtrlWindow2", winProp)
}
function chkWin() { if (OpenWin) { OpenWin.focus(); } }
function clearWin() { OpenWin = null; } 


function chkNum(field, Min)
{
	if (field.value != '')
	{
		if (!MyIsNumeric(getNumeric(field.value)))
		{
			alert(txtValNumVal)
			field.value = ""
			field.focus()
		}
		else if (Min != '')
		{
			if (parseFloat(getNumeric(field.value)) < paserFloat(Min))
			{
				alert(txtValNumMinVal.replace("{0}", Min));
				field.value = "";
				field.focus();
			}
		}
	}
}

function changeBranch(LineID, Value)
{
	if (Value.indexOf('*') != -1)
	{
		var BankCode;
		if (LineID == "new")
		{
			BankCode = document.Form1.banco.value;
		}
		else
		{
			BankCode = document.getElementById("banco" + LineID).value;
		}
		if (BankCode != '')
		{
			varLineID = LineID;
			Start2("branchs.asp?BankCode="+BankCode+"&Value="+Value+"&pop=Y&rParent=Y&AddPath=../","272","200","");
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


function chkCheck(Field)
{
	if (Field.value != '')
	{
		if (MyIsNumeric(getNumeric(Field.value)))
		{
			if (parseFloat(getNumeric(Field.value)) > 2147483647)
			{
				Field.value = 2147483647;
			}
			else if (parseFloat(getNumeric(Field.value)) < -2147483648)
			{
				Field.value = -2147483648;
			}
		}
	}
}

function delCheck(lineID)
{
	if(confirm(txtConfDelChk.replace('{0}', lineID+1)))doMyLink('submit.asp', 'submitCmd=payCheck&delete.x=y&lineNum=' + lineID, '');
}

function SetSaldo()
{
	var openBal = ClearCur(document.Form1.ImpInc.value)-ClearCur(document.Form1.Pagado.value);
	document.Form1.SaldoPag.value = DocCur + ' ' + OLKFormatNumber(openBal,SumDec);
	document.Form1.pagVal.value = document.Form1.Pagado.value;
}

function chkNewChk()
{
	if (document.Form1.impval.value != '')
	{
		document.Form1.Agregar.value = "Y";
		document.Form1.submit() ;
	}
	else
	{
		alert(txtValChkImp);
	}
}


function setTotal(Field, Max)
{
	var varTotal = 0;
	var LineNum;
	if (document.Form1.LineNum)
	{
		if (document.Form1.LineNum.length)
		{
			for (var i = 0;i<document.Form1.LineNum.length;i++)
			{
				LineNum = document.Form1.LineNum[i].value;
				varTotal += ClearCur(document.getElementById('imp' + LineNum).value);
			}
		}
		else
		{
			LineNum = document.Form1.LineNum.value;
			varTotal = ClearCur(document.getElementById('imp' + LineNum).value);
		}
	}
	
	if (document.Form1.impval.value != '') varTotal += ClearCur(document.Form1.impval.value);
	document.Form1.Total.value = DocCur + ' ' + OLKFormatNumber(varTotal,SumDec);
	document.Form1.Pagado.value = DocCur + ' ' + OLKFormatNumber(getNumeric(vPagado)-parseFloat(getNumeric(checkSum))+varTotal,SumDec);
	var openBal = ClearCur(document.Form1.ImpInc.value)-ClearCur(document.Form1.Pagado.value);;
	document.Form1.SaldoPag.value = DocCur + ' ' + OLKFormatNumber(openBal,SumDec);
	document.Form1.pagVal.value = document.Form1.Pagado.value;
	if (Field.value != '') Field.value = DocCur + ' ' + OLKFormatNumber(ClearCur(Field.value),SumDec);
}
