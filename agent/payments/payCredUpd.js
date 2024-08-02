function setCard(CreditCard, CardName, CrTypeCode, CrTypeName, AcctCode, MinCreditVal, MinToPayVal, MaxValidVal, InstalMentVal, AcctDisp)
{
	document.Form3.tarjeta.value = CardName;
	document.Form3.CreditAcct.value = AcctCode;
	document.Form3.CreditAcctDisp.value = AcctDisp;
	document.Form3.CreditCard.value = CreditCard;
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
	if (document.Form3.SistPagCode.value == 0)
	{
		alert(txtValPymntSys);
		return false;
	}
	else if (document.Form3.impval.value == "")
	{
		alert(txtValCardNoPymnt);
		return false;
	}
	else if (document.Form3.CardValidM.value == "" || document.Form3.CardValidY.value == "")
	{
		alert(txtValCardDueDat);
		return false;
	}
	else if (MaxValid != "" && MaxValid != null)
	{
		if (parseFloat(getNumeric(document.Form3.impval.value)) > parseFloat(getNumeric(MaxValid)))
		{
			if (document.Form3.autorizacion.value == "")
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
	document.Form3.SistPagCode.value = CrTypeCode;
	document.Form3.syspag.value = CrTypeName;
	
	if (InstalMentVal == "Y")
	{
		document.Form3.pagcant.readonly = false;	
		document.Form3.pagcant.className = "";
		document.Form3.perpago.readonly = false;		
		document.Form3.perpago.className= "";
		document.Form3.cpagoadd.readonly = false;
		document.Form3.cpagoadd.className="";
	}
	else
	{
		document.Form3.pagcant.readonly = true;
		document.Form3.pagcant.className= "InputDes";
		document.Form3.pagcant.value = "1";
		document.Form3.perpago.readonly = true;		
		document.Form3.perpago.className= "InputDes";
		document.Form3.cpagoadd.readonly = true;
		document.Form3.cpagoadd.className= "InputDes";
	}
	if (document.Form3.impval.value != "") optChangeImp();
}
				
function chkVal(field)
{
	if (field.name == "pagcant")
	{
		if (parseFloat(getNumeric(field.value)) <= 0) field.value = 0;
	}
	else if (field.name == "impval" && field.value != "")
	{
		if (ClearCur(field.value) > (ClearCur(document.Form2.SaldoPag.value)+parseFloat(getNumeric(creditsum))))
		{
			field.value = ClearCur(document.Form2.SaldoPag.value)+parseFloat(getNumeric(creditsum));
		}
		else if (ClearCur(field.value) <= 0)
		{
			field.value = 0;
		}
		else if (MinCredit != "" && MinCredit != null)
		{
			if (ClearCur(field.value) < parseFloat(getNumeric(MinCredit)))
			{
				field.value = MinCredit;
				alert(txtMinCredPymnt.replace("{0}", MinCredit));
			}
		}
		field.value = DocCur + ' ' + OLKFormatNumber(ClearCur(field.value),SumDec)
	}
	else if (field.name == "perpago" ||  field.name == "cpagoadd")
	{
		if (field.value != "")
		{
			if (parseFloat(getNumeric(field.value)) < 0) optChangeImp("Y");
			field.value = DocCur + ' ' + OLKFormatNumber(ClearCur(field.value),SumDec);
		}
		else
		{
			switch (field.name)
			{
				case "perpago":
					optChangeCPago();
					break;
				case "cpagoadd":
					optChangePerPago();
					break;
			}
		}
	}
}

function optChangeImp()
{
	if (!document.Form3.perpago.disabled && document.Form3.impval.value != '')
	{
		var2 = ClearCur(document.Form3.impval.value)/ClearCur(document.Form3.pagcant.value);
		var1 = ClearCur(document.Form3.impval.value)-(var2*(ClearCur(document.Form3.pagcant.value)-1));
		document.Form3.perpago.value = DocCur + ' ' + OLKFormatNumber(var1,SumDec);
		if (parseFloat(getNumeric(document.Form3.pagcant.value)) > 1)
		{
			document.Form3.cpagoadd.value = "";
			document.Form3.cpagoadd.readOnly = false;
			document.Form3.cpagoadd.className = "";
			document.Form3.perpago.readOnly = false;
			document.Form3.perpago.className = "";
			document.Form3.cpagoadd.value = DocCur + ' ' + OLKFormatNumber(var2,SumDec);
		}
		else
		{
			document.Form3.cpagoadd.readOnly = true;
			document.Form3.cpagoadd.className = "InputDes";
			document.Form3.perpago.readOnly = true;
			document.Form3.perpago.className = "InputDes";
			document.Form3.cpagoadd.value = "";
		}
	}
}

function optChangePerPago()
{
	var1 = ClearCur(document.Form3.perpago.value);
	var2 = (ClearCur(document.Form3.impval.value)-var1)/(ClearCur(document.Form3.pagcant.value)-1);
	document.Form3.cpagoadd.value = DocCur + ' ' + OLKFormatNumber(var2,SumDec);
}
function optChangeCPago()
{
	var2 = ClearCur(document.Form3.cpagoadd.value)*(ClearCur(document.Form3.pagcant.value)-1);
	var1 = ClearCur(document.Form3.impval.value)-var2;
	document.Form3.perpago.value = DocCur + ' ' + OLKFormatNumber(var1,SumDec);
}
function StartSisPag()
{
	if (!SPago)
	{
		Start("cardsSist.asp","200","200","auto");
	}
}