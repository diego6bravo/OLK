function replaceSubstring(inputString, fromString, toString) {
   // Goes through the inputString and replaces every occurrence of fromString with toString
   var temp = inputString;
   if (fromString == "") {
      return inputString;
   }
   if (toString.indexOf(fromString) == -1) { // If the string being replaced is not a part of the replacement string (normal situation)
      while (temp.indexOf(fromString) != -1) {
         var toTheLeft = temp.substring(0, temp.indexOf(fromString));
         var toTheRight = temp.substring(temp.indexOf(fromString)+fromString.length, temp.length);
         temp = toTheLeft + toString + toTheRight;
      }
   } else { // String being replaced is part of replacement string (like "+" being replaced with "++") - prevent an infinite loop
      var midStrings = new Array("~", "`", "_", "^", "#");
      var midStringLen = 1;
      var midString = "";
      // Find a string that doesn't exist in the inputString to be used
      // as an "inbetween" string
      while (midString == "") {
         for (var i=0; i < midStrings.length; i++) {
            var tempMidString = "";
            for (var j=0; j < midStringLen; j++) { tempMidString += midStrings[i]; }
            if (fromString.indexOf(tempMidString) == -1) {
               midString = tempMidString;
               i = midStrings.length + 1;
            }
         }
      } // Keep on going until we build an "inbetween" string that doesn't exist
      // Now go through and do two replaces - first, replace the "fromString" with the "inbetween" string
      while (temp.indexOf(fromString) != -1) {
         var toTheLeft = temp.substring(0, temp.indexOf(fromString));
         var toTheRight = temp.substring(temp.indexOf(fromString)+fromString.length, temp.length);
         temp = toTheLeft + midString + toTheRight;
      }
      // Next, replace the "inbetween" string with the "toString"
      while (temp.indexOf(midString) != -1) {
         var toTheLeft = temp.substring(0, temp.indexOf(midString));
         var toTheRight = temp.substring(temp.indexOf(midString)+midString.length, temp.length);
         temp = toTheLeft + toString + toTheRight;
      }
   } // Ends the check to see if the string being replaced is part of the replacement string or not
   return temp; // Send the updated string back to the user
} // Ends the "replaceSubstring" function

function IsNumeric(sText)
{
   var ValidChars = '0123456789' + formatDec;
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

function valDelItems()
{
	if (document.frmCart.chkDel)
	{
		var found = 0;
		if (document.frmCart.chkDel.length)
		{
			for (var i = 0;i<document.frmCart.chkDel.length;i++)
			{
				found += document.frmCart.chkDel[i].checked ? 1 : 0;
			}
		}
		else
			found = document.frmCart.chkDel.checked ? 1 : 0;
			
		if (found)
			return confirm(txtConfDelItm.replace('{0}', found));
		else
			return false;
	}
	else
		return false;
}

function FormatNumber(value, decimals) 
{
	return formatNumberDec(value, decimals, false);
}

function chkNum(Field, Max, Dec) 
{
	if (!IsNumeric(Field.value)) 
	{
		Field.value = Max
		alert(txtValNumVal) 
	}
	Field.value = FormatNumber(Field.value,Dec)
	
}

function setDocDisc()
{
	chkNum(document.frmCart.DiscPrcnt, document.frmCart.oldDiscPrcnt.value, PercentDec);
	
	if (parseFloat(document.frmCart.DiscPrcnt.value) > MaxDocDisc)
	{
		alert(txtMaxDiscount.replace('{0}', MaxDocDisc));
		document.frmCart.DiscPrcnt.value = MaxDocDisc;
	}
	
	document.frmCart.oldDiscPrcnt.value = document.frmCart.DiscPrcnt.value;
	
	setTotal();
}


function chkPrice(Field, RestVal, Dec, ManPrc, LineTotal, Qty, SalPackUn, un, Currency, UnitPrice)
{
	chkNum(Field, RestVal, Dec);
	
	var priceVal = parseFloat(Field.value.replace(formatDec, '.'));
	var unitPriceVal = parseFloat(UnitPrice.value.replace(formatDec, '.'));
	
	if (priceVal > 8699999999999.000)
	{
		alert(txtValNumMaxVal.replace('{0}', 8699999999999.000));
		Field.value = FormatNumber(8699999999999.000, Dec)
	}
	else
	{
		var discVal = 100-(100*priceVal/unitPriceVal);
		var maxDiscVal = parseFloat(MaxDiscount.replace(formatDec, '.'));
		
		if (discVal > maxDiscVal)
		{
			alert(txtMaxDiscount.replace('{0}', MaxDiscount));
			Field.value = FormatNumber(unitPriceVal-(unitPriceVal*(maxDiscVal/100)), Dec);
		}
	}

	ManPrc.value='Y';
	updateLineTotal(LineTotal, Field.value, Qty.value, SalPackUn.value, un.value, Currency.value)
}
function chkQty(Field, Max, Dec, Item, ManPrc, LineTotal, Price, NumInSale, SalPackUn, Unit, Currency, SetPrice) 
{
	chkNum(Field, Max, Dec);
	
	if (parseFloat(Field.value.replace(formatDec, '.')) > 32999999999.00)
	{
		alert(txtValNumMaxVal.replace('{0}', 32999999999.00));
		Field.value = FormatNumber(32999999999.00, Dec)
	}
	
	if (itemVolDisc != null && ManPrc.value == 'N')
	{
		var arrVolDisc = itemVolDisc.split('{I}');
		
		for (var i = 0;i<arrVolDisc.length;i++)
		{
			var data = arrVolDisc[i].split('{D}');
			if (data[0] == Item)
			{
				var setPriceBy = 1;
				if (parseInt(Unit) > 1) setPriceBy = parseFloat(NumInSale);
				if (!UnEmbPriceSet && parseInt(Unit) == 3) setPriceBy=setPriceBy* parseFloat(SalPackUn);
				var foundDisc = false;
				var arrData = data[1].split('{S}');
				for (var i = arrData.length-1;i>=0;i--)
				{
					var volDiscItm = arrData[i].split('|');
					if (parseFloat(Field.value) >= parseFloat(volDiscItm[0]))
					{
						Price.value = FormatNumber(volDiscItm[1]*setPriceBy, PriceDec);
						foundDisc = true;
						break;
					}
				}
				
				if (!foundDisc)
				{
					Price.value = FormatNumber(SetPrice*setPriceBy, PriceDec);
				}
				break;
			}
		}
	}
	
	updateLineTotal(LineTotal, Price.value, Field.value, SalPackUn, Unit, Currency);
}


function updateLineTotal(Field, Price, Quantity, SalPackUn, UnEmb, Currency) 
{
	if (UnEmb == 3 && UnEmbPriceSet) 
	{ 
		Field.value = Currency + ' ' + FormatNumber(ConvertToFloat(Price)*ConvertToFloat(Quantity)*SalPackUn,SumDec) 
	}
	else 
	{
		Field.value = Currency + ' ' + FormatNumber(ConvertToFloat(Price)*ConvertToFloat(Quantity),SumDec) 
	}
	setTotal() 
}


function ConvertToFloat(value)
{
	return parseFloat(value.replace(formatDec, '.'));
}

function changeUnEmb(NewUnEmb, Price, Cant, UnEmb, LineTotal, SalPackUn, NumInSale, Currency) 
{
	if (UnEmb.value == 2 && NewUnEmb.value == 1) { Price.value = FormatNumber(ConvertToFloat(Price.value)/NumInSale,PriceDec) }
	else if (UnEmb.value == 1 && NewUnEmb.value == 2) { Price.value = FormatNumber(ConvertToFloat(Price.value)*NumInSale,PriceDec) }
	else if (UnEmb.value == 3 && NewUnEmb.value == 2) { Price.value = FormatNumber(ConvertToFloat(Price.value)/(!UnEmbPriceSet ? SalPackUn : 1),PriceDec) }
	else if (UnEmb.value == 2 && NewUnEmb.value == 3) { Price.value = FormatNumber(ConvertToFloat(Price.value)*(!UnEmbPriceSet ? SalPackUn : 1),PriceDec) }
	else if (UnEmb.value == 3 && NewUnEmb.value == 1) { Price.value = FormatNumber((ConvertToFloat(Price.value)/(!UnEmbPriceSet ? SalPackUn : 1))/NumInSale,PriceDec) }
	else if (UnEmb.value == 1 && NewUnEmb.value == 3) { Price.value = FormatNumber((ConvertToFloat(Price.value)*NumInSale)*(!UnEmbPriceSet ? SalPackUn : 1),PriceDec) }
	UnEmb.value = NewUnEmb.value
	updateLineTotal(LineTotal, Price.value, Cant.value, SalPackUn, NewUnEmb.value, Currency)
}

function valFrm()
{
	if (document.frmCart.VerfyDueDate.value == 'N')
	{
		alert(txtValDelDate);
		return false;
	}
	return true;
}

function resetFastAdd()
{
	document.frmAddFast.SaleUnit.value = AgentSaleUnit;
	document.frmAddFast.txtFastAddQty.value = FormatNumber(1, QtyDec); 
	document.frmAddFast.txtFastAdd.value = '';
	document.frmAddFast.txtFastAdd.focus();
}

function valFastAdd()
{
	if (document.frmAddFast.txtFastAdd.value == '')
	{
		alert(txtValEnterValue);
		document.frmAddFast.txtFastAdd.focus();
		return false;
	}
	
	if (document.frmAddFast.txtFastAddQty.value == '')
	{
		alert(txtValEnterValue);
		document.frmAddFast.txtFastAddQty.focus();
		return false;
	}
	
	if (!IsNumeric(document.frmAddFast.txtFastAddQty.value))
	{
		alert(txtValNumVal);
		document.frmAddFast.txtFastAddQty.value = '';
		document.frmAddFast.txtFastAddQty.focus();
		return false;
	}
	
	if (parseFloat(document.frmAddFast.txtFastAddQty.value.replace(formatDec, '')) > 32999999999.00)
	{
		alert(txtValNumMaxVal.replace('{0}', '32,999,999,999.00'));
		document.frmAddFast.txtFastAddQty.value = '';
		document.frmAddFast.txtFastAddQty.focus();
		return false;
	}

	return true;
}

function onScan(ev){
var scan = ev.data;
	document.frmAddFast.txtFastAdd.value = scan.value;
	document.frmAddFast.submit();
}
function onSwipe(ev){
}

try
{
document.addEventListener("BarcodeScanned", onScan, false);
document.addEventListener("MagCardSwiped", onSwipe, false);
}
catch(err) {}
