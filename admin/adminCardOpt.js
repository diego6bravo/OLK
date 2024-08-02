var rsHasVars = false;

function valFrm()
{
	rowName = document.form1.rowName;
	for (var i = 0;i<rowName.length;i++)
	{
		if (rowName[i].value == '')
		{
			alert(LtxtValFldNam);
			rowName[i].focus();
			return false;
		}
	}
	return true;
}

function valFrm2()
{
	if (document.form2.valQuery.value == 'Y')
	{
		alert(LtxtValQryVal);
		document.form2.btnVerfyFilter.focus();
		return false;
	}
	else if (document.form2.rowName.value == '')
	{
		alert(LtxtValFldNam2);
		document.form2.rowName.focus();
		return false;
	}
	else if (document.form2.customSql.value == '')
	{
		alert(LtxtValQry);
		document.form2.customSql.focus();
		return false;
	}
	return true;
}

function IsNumeric(sText)
{
   var ValidChars = "0123456789.";
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

var myBtnVerfy;
function VerfyQuery()
{
	document.frmVerfyQuery.Query.value = document.form2.customSql.value;
	myBtnVerfy = document.form2.btnVerfy;
	document.frmVerfyQuery.submit();
}

function VerfyQueryVerified()
{
	//myBtnVerfy.disabled = true;
	document.form2.btnVerfyFilter.src='images/btnValidateDis.gif'
	document.form2.btnVerfyFilter.style.cursor = '';
	document.form2.valQuery.value='N';
}