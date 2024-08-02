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

function FormatNumber(expr, decplaces) 
{
	return formatNumberDec(expr, decplaces, false);
}
function setLineTotal() 
{
	if (SaleType == 3) 
	{
		document.addcart.LineTotal.value = FormatNumber(parseFloat(document.addcart.Quantity.value)*parseFloat(document.addcart.precio.value)*(UnEmbPriceSet ? SalPackUn : 1),SumDec) 
	}
	else 
	{
		document.addcart.LineTotal.value = FormatNumber(parseFloat(document.addcart.Quantity.value)*parseFloat(document.addcart.precio.value),SumDec) 
	}
}
function chkThis(Field, Old, Dec, Max)
{
	if (!IsNumeric(Field.value, false) || Field.value < 0 || Field.value == '') 
	{ 
		Field.value = FormatNumber(Old.value, Dec); 
		return;
	}
	
	if (parseFloat(Field.value) > Max)
	{
		alert(txtValNumMaxVal.replace('{0}', Max));
		Field.value = FormatNumber(Old.value, Dec); 
		return;
	}
	Field.value = FormatNumber(Field.value, Dec);
	Old.value = Field.value;
	
	switch(Field.name)
	{
		case 'precio':
			UnitPrice = parseFloat(document.addcart.UnitPrice.value);
			Price = parseFloat(Field.value);
			
			if (UnitPrice != 0)
			{
				SaleType = parseInt(document.addcart.SaleType.value);
				switch (SaleType)
				{
					case 2:
						UnitPrice = UnitPrice*NumInSale;
						break;
					case 3:
						UnitPrice = UnitPrice*NumInSale*(!UnEmbPriceSet ? SalPackUn : 1);
						break;
				}
				document.addcart.DiscPrcnt.value = FormatNumber(100-(Price*100)/UnitPrice, PercentDec);
			}
			else
			{
				document.addcart.DiscPrcnt.value = FormatNumber(0, PercentDec);
			}
			changePrice = false;
			chkDiscount(document.addcart.DiscPrcnt);
			changePrice = true;
			
			document.addcart.ManPrc.value = 'Y';
			break;
		case 'Quantity':
			if (document.addcart.ManPrc.value == 'N')
			{
				if (itemVolDisc != null)
				{
					var setPriceBy = 1;
					if (parseInt(document.addcart.SaleType2.value) > 1) setPriceBy = NumInSale;
					if (UnEmbPriceSet) { if (parseInt(document.addcart.SaleType2.value) == 3) setPriceBy = setPriceBy * SalPackUn; }
					var foundDisc = false;
					var arrVolDisc = itemVolDisc.split('{S}');
					for (var i = arrVolDisc.length-1;i>=0;i--)
					{
						var volDiscItm = arrVolDisc[i].split('|');
						if (parseFloat(Field.value) >= parseFloat(volDiscItm[0]))
						{
							document.addcart.precio.value = FormatNumber(volDiscItm[1]*setPriceBy, PriceDec);
							foundDisc = true;
							break;
						}
					}
					if (!foundDisc)
					{
						document.addcart.precio.value = FormatNumber(varPrice*setPriceBy, PriceDec);
					}
				}
			}
			break;
	}
	
	setLineTotal()
}

function chkDiscount(Field)
{
	chkVal = Field.value;
	if (chkVal == '')
	{
		Field.value = FormatNumber(document.addcart.oldDiscPrcnt.value, PercentDec);
		return;
	}
	
	if (!IsNumeric(chkVal, true))
	{
		Field.value = FormatNumber(document.addcart.oldDiscPrcnt.value, PercentDec);
		return;
	}
	
	if (parseFloat(chkVal) > parseFloat(MaxDiscount))
	{
		alert(txtMaxDiscount.replace('{0}', MaxDiscount));
		chkVal = MaxDiscount;
		changePrice = true;
	}
	
	Field.value = chkVal;
	if (changePrice) setDiscPrice();
	
	Field.value = FormatNumber(parseFloat(Field.value), PercentDec);
}

function setDiscPrice()
{
	var disc = parseFloat(document.addcart.DiscPrcnt.value);
	
	document.addcart.oldDiscPrcnt.value = disc;
	document.addcart.ManPrc.value = 'Y';
	
	UnitPrice = parseFloat(document.addcart.UnitPrice.value);
	if (UnitPrice != 0)
	{
		SaleType = parseInt(document.addcart.SaleType.value);
		switch (SaleType)
		{
			case 2:
				UnitPrice = UnitPrice*NumInSale;
				break;
			case 3:
				UnitPrice = UnitPrice*NumInSale*(!UnEmbPriceSet ? SalPackUn : 1);
				break;
		}
		document.addcart.precio.value = FormatNumber(UnitPrice-((disc*UnitPrice)/100), PriceDec)
	}
	else
	{
		document.addcart.precio.value = FormatNumber(0, PriceDec);
	}
	document.addcart.oldPrecio.value = document.addcart.precio.value;
	setLineTotal();
}

function changeUnEmb(NewUnEmb, Price) 
{
	if ((SaleType == 2) && (NewUnEmb.value == 1)) { Price.value = FormatNumber(parseFloat(Price.value)/NumInSale,PriceDec); } 
	else if ((SaleType == 1) && (NewUnEmb.value = 2)) { Price.value = FormatNumber(parseFloat(Price.value)*NumInSale,PriceDec); }
	else if ((SaleType == 2) && (NewUnEmb.value = 3)) { Price.value = FormatNumber(parseFloat(Price.value)*(!UnEmbPriceSet ? SalPackUn : 1),PriceDec) }
	else if ((SaleType == 3) && (NewUnEmb.value = 2)) { Price.value = FormatNumber(parseFloat(Price.value)/(!UnEmbPriceSet ? SalPackUn : 1),PriceDec) }
	else if ((SaleType == 1) && (NewUnEmb.value = 3)) { Price.value = FormatNumber(parseFloat(Price.value)*NumInSale*(UnEmbPriceSet ? SalPackUn : 1),PriceDec) }
	else if ((SaleType == 3) && (NewUnEmb.value = 1)) { Price.value = FormatNumber(parseFloat(Price.value)/(!UnEmbPriceSet ? SalPackUn : 1)/NumInSale,PriceDec) }
	SaleType = NewUnEmb.value
	setLineTotal()
}
