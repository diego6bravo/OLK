function setCard(CreditCard, CardName, CrTypeCode, CrTypeName, AcctCode, MinCreditVal, MinToPayVal, MaxValidVal, InstalMentVal, AcctDisp)
{
	document.Form1.tarjeta.value = CardName;
	document.Form1.CreditAcct.value = AcctCode;
	document.Form1.CreditAcctDisp.value = AcctDisp;
	document.Form1.CreditCard.value = CreditCard;
	MinCredit = MinCreditVal;
	MinToPay = MinToPayVal;
	MaxValid = MaxValidVal;
	InstalMent = InstalMentVal;
	if (CrTypeCode != "")
	{
		setSistPag(CrTypeCode, CrTypeName, MinCreditVal, MinToPayVal, MaxValidVal, InstalMentVal);
		SPago = true;
	}
	else
	{
		SPago = false;
	}
}
			

function valFrm()
{
	if (document.Form1.SistPagCode.value == 0)
	{
		alert(txtValPymntSys);
		return false;
	}
	else if (document.Form1.impval.value == "")
	{
		alert(txtValCardNoPymnt);
		return false;
	}
	else if (document.Form1.CardValidM.value == "" || document.Form1.CardValidY.value == "" || document.Form1.CardValidM.value == "MM" || document.Form1.CardValidY.value == "YY")
	{
		alert(txtValCardDueDat);
		return false;
	}
	else if (MaxValid != "" && MaxValid != null)
	{
		if (ClearCur(document.Form1.impval.value) > MaxValid)
		{
			if (document.Form1.autorizacion.value == "")
			{
				alert(txtValAut);
				return false;
			}
		}
	}
	return true;
}
		
		
function setSistPag(CrTypeCode, CrTypeName, MinCreditVal, MinToPayVal, MaxValidVal, InstalMentVal)
{
	document.Form1.SistPagCode.value = CrTypeCode;
	document.Form1.syspag.value = CrTypeName;
	if (InstalMentVal == "Y")
	{
		document.Form1.pagcant.readonly = false;
		document.Form1.pagcant.className = "";
	}
	else
	{
		document.Form1.pagcant.readonly = true;
		document.Form1.pagcant.className = "InputDes";
		document.Form1.pagcant.value = "1";
	}
	if (document.Form1.impval.value != "") optChangeImp("N");
}
				
function chkVal(field)
{
	if (field.name == "pagcant")
	{
		if (parseFloat(getNumeric(field.value)) <= 0) field.value = 1
	}
	else if (field.name == "impval" && field.value != "")
	{
		if (ClearCur(field.value) <= 0)
		{
			field.value = 0;
		}
		else if (MinCredit != "" && MinCredit != null)
		{
			if (ClearCur(field.value) < ClearCur(MinCredit))
			{
				field.value = MinCredit;
				alert(txtMinCredPymnt.replace("{0}", MinCredit));
			}
		}
		field.value = DocCur + ' ' + OLKFormatNumber(getNumeric(field.value),SumDec)
	}
	else if (field.name == "perpago" || field.name == "cpagoadd")
	{
		if (ClearCur(field.value) < 0) optChangeImp("Y");
		field.value = DocCur + ' ' + OLKFormatNumber(getNumeric(field.value),SumDec)
	}
}

function optChangeImp(ChgPag)
{
	if (parseFloat(getNumeric(document.Form1.pagcant.value)) > 1)
	{
			document.Form1.perpago.readonly = false;	
			document.Form1.perpago.className = "";
			document.Form1.cpagoadd.readonly = false;
			document.Form1.cpagoadd.className = "";
	}
	else
	{
			document.Form1.perpago.readonly = true;	
			document.Form1.perpago.className = "InputDes";
			document.Form1.cpagoadd.readonly = true;
			document.Form1.cpagoadd.className = "InputDes";
	}
	if (!document.Form1.perpago.disabled && document.Form1.impval.value != "")
	{
		var2 = ClearCur(document.Form1.impval.value)/ClearCur(document.Form1.pagcant.value);
		var1 = ClearCur(document.Form1.impval.value)-(var2*(ClearCur(document.Form1.pagcant.value)-1));
		document.Form1.perpago.value = DocCur + ' ' + OLKFormatNumber(var1,SumDec);
		if (document.Form1.pagcant.value > 1)
		{
			document.Form1.cpagoadd.value = DocCur + ' ' + OLKFormatNumber(var2,SumDec);
		}
		else
		{
			document.Form1.cpagoadd.value = "";
		}
	}
}

function optChangePerPago()
{
	var1 = ClearCur(document.Form1.perpago.value);
	var2 = (ClearCur(document.Form1.impval.value)-var1)/(ClearCur(document.Form1.pagcant.value)-1);	
	document.Form1.cpagoadd.value = DocCur + ' ' + OLKFormatNumber(var2,SumDec);
}

function optChangeCPago()
{
	var2 = ClearCur(document.Form1.cpagoadd.value)*(ClearCur(document.Form1.pagcant.value)-1);
	var1 = ClearCur(document.Form1.impval.value)-var2;
	document.Form1.perpago.value = DocCur + ' ' + OLKFormatNumber(var1,SumDec);
}

function StartSisPag()
{
	if (!SPago)
	{
		Start("cardsSist.asp","200","200","auto");
	}
}