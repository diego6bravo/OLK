function goEditItem(item)
{
	doMyLink('addItem/goEditItem.asp', 'AddPath=&ItemCode=' + item, '');
}
function goDupItem(item)
{
	doMyLink('addItem/goEditItem.asp', 'AddPath=&Duplicate=Y&ItemCode=' + item, '')
}
function Start(popType, page) {

	
	if (UserType == 'V')
	{
		if (popType == 'W') popH = 420;
	}
	if (popType != 'Rec')
	{
		switch (UserType)
		{
			case 'C':
				OpenWin = this.open(page, "searchCart", "toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no,width=482,height=" + popH);
				break;
			case 'V':
				OpenWin = this.open(page, "searchCart", "toolbar=no,menubar=no,location=no,resizable=no,scrollbars=yes,width=598,height=" + popH + ",status=yes");
				break;
		}
	}
	else
	{
		OpenWin = this.open(page, "searchCart", "toolbar=no,menubar=no,location=no,resizable=no,scrollbars=yes,width=598,height=420,status=yes");
	}
	OpenWin.focus()
}
function GoStatus() { }

function goAddChkItems()
{
	var items = '';
	var qtys = '';
	var openRec = false;
	if (chkItem.length)
	{
		for (var i = 0;i<chkItem.length;i++)
		{
			if (chkItem[i].checked)
			{
				if (items != '') { items += '{S}'; qtys += '{S}'; }
				
				items += 'N\'' + chkItem[i].value + '\'';
				if (Qty) qtys += Qty[i].value; else qtys += '1';
				
				if (TreeType[i].value == 'T')
				{
					openRec = true;
				}
			}
		}
	}
	else
	{
		if (chkItem.checked)
		{
			items = 'N\'' + chkItem.value + '\'';
			if (Qty) qtys += Qty.value; else qtys += '1';
			
			if (TreeType.value == 'T')
			{
				openRec = true;
			}
		}
	}
	
	if (items != '')
	{
		if (!openRec)
		{
			document.frmGoAddItem.action = 'cart/addCartSubmitMulti.asp';
			document.frmGoAddItem.Item.value = items;
			document.frmGoAddItem.T1.value = qtys;
			document.frmGoAddItem.submit();
		}
		else
		{
			document.frmGoItem.action = 'cart/addCartRec.asp';
			document.frmGoItem.Item.value = items;
			document.frmGoItem.T1.value = qtys;
			Start('Rec', '');
			document.frmGoItem.submit();
		}
	}
	else
	{
		alert(txtChkAtLead1Item);
	}
}
function chkAllItems(btnChk)
{
	if (chkItem.length)
	{
		for (var i = 0;i<chkItem.length;i++)
		{
			chkItem[i].checked = btnChk.checked;
		}
	}
	else
	{
		chkItem.checked = btnChk.checked;
	}
	if (chkAllItms[0] != btnChk) chkAllItms[0].checked = btnChk.checked; 
	if (chkAllItms[1] != btnChk) chkAllItms[1].checked = btnChk.checked; 
}

function chkNum(fld, bookmark)
{
	if (chkItem.length)
	{
		chkItem[bookmark-1].checked = true;
	}
	else 
	{
		chkItem.checked = true;
	}
}


function chkItemChkAll()
{
	var chk = true;
	if (chkItem.length)
	{
		for (var i = 0;i<chkItem.length;i++)
		{
			if (!chkItem[i].checked)
			{
				chk = false;
				break;
			}
		}
	}
	else
	{
		chk = chkItem.checked;
	}
	chkAllItms[0].checked = chk; 
	chkAllItms[1].checked = chk; 
}
function goViewItem(Item) 
{ 
	switch (UserType)
	{
		case 'C':
			document.frmGoItem.action = 'item.asp';
			document.frmGoItem.Item.value = Item;
			document.frmGoItem.cmd.value = 'a';
			document.frmGoItem.submit();
			break;
		case 'V':
			ItemCmd = searchCmd == 'searchCatalog' ? 'D' : 'A';
			openItemDetails(Item);	
			break;
	}
}


function goWishList(Item, PackPrice)
{	switch (UserType)
	{
		case 'C':
			document.frmGoItem.action = 'item.asp';
			document.frmGoItem.Item.value = Item;
			document.frmGoItem.cmd.value = 'w';
			document.frmGoItem.submit();
			break;
		case 'V':
			ItemCmd = 'W';
			openItemDetails(Item);	
			break;
	}
}
var qtyVal = 1;
function goAddItemQty(Item, TreeType, bookmark)
{
	if (!isAnon)
	{
	
		if (Qty) 
		{
			qtyVal = Qty.length ? Qty[parseInt(bookmark)-1].value : Qty.value; 
		}
		else 
			qtyVal = 1;
	
		goAddItem(Item, TreeType);
	}
	else
	{
		alert(txtStartSesion);
	}
}

function goAddItem(Item, TreeType)
{
	if (TreeType != 'T')
	{
		document.frmGoAddItem.action = 'cart/addCartSubmitM.asp';
		document.frmGoAddItem.Item.value = Item;
		document.frmGoAddItem.T1.value = qtyVal;
		setFlowAlertVars('D2', (Item + '{S}' + qtyVal + '{S}{S}{S}'), 'document.frmGoAddItem.DocConf.value=typeIDs;document.frmGoAddItem.submit();', '');
		doFlowAlert();

	}
	else
	{
		document.frmGoItem.action = 'cart/addCartRec.asp';
		document.frmGoItem.Item.value = 'N\'' + Item + '\'';
		document.frmGoItem.T1.value = qtyVal;
		Start('Rec', '');
		document.frmGoItem.submit();
	}
}

function goPage(p) { document.frmGPage.page.value = p; document.frmGPage.submit(); }
function printCat(pdf) 
{
	oldAction = document.frmGPage.action;
	if (pdf == 'N')	document.frmGPage.action = "searchCartPDF.asp";
	else document.frmGPage.action = "makePdf.asp";
	document.frmGPage.target = "_blank";
	document.frmGPage.PrintCatalog.value = "Y";
	document.frmGPage.submit();
	document.frmGPage.PrintCatalog.value = "";
	document.frmGPage.target = "";
	document.frmGPage.action = oldAction;
}
