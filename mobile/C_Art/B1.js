function valFrm()
{
	if (document.frmAddItems.chkItem.length)
	{
		var found = false;
		for (var i = 0;i<document.frmAddItems.chkItem.length;i++)
		{
			if (document.frmAddItems.chkItem[i].checked)
			{
				found = true;
				break;
			}
		}
		if (!found)
		{
			alert(txtChkAtLead1Item);
			return false;
		}
	}
	else
	{
		if (!document.frmAddItems.chkItem.checked)
		{
			alert(txtChkAtLead1Item);
			return false;
		}
	}	
	return true;
}

function goP(p)
{
	window.location.href='operaciones.asp?cmd=searchItems&slist=' + slist + '&page=' + p;
}

function IsNumeric(sText, Negative)
{
   var ValidChars = '0123456789' + GetFormatDec;
   if (Negative) ValidChars += '-';
   var IsNumber=true;
   var Char;

 
   for (i = 0; i < sText.length && IsNumber == true; i++) 
      { 
      Char = sText.charAt(i); 
      if (ValidChars.indexOf(Char) == -1) 
         {
         IsNumber = false;
         }
      }
   return IsNumber;
   
   }

function chkQty(fld, o, d, m)
{
	if (!IsNumeric(fld.value, false) || fld.value == '') 
	{ 
		fld.value = o; 
		return;
	}
	
	if (parseFloat(fld.value) > m)
	{
		alert(txtValNumMaxVal.replace('{0}', m));
		fld.value = o;
		return;
	}
	fld.value = FormatNumber(fld.value, d);

}

function FormatNumber(expr, decplaces) 
{
	return formatNumberDec(exprm, decplaces, false);
}
