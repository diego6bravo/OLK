function doDetail(ObjectCode, Entry, linkCat, IsEntry, high, linkPop)
{
	document.frmViewDetail.target = linkCat == 'N' ? '_blank' : '';
	document.frmViewDetail.AddPath.value = '';
	document.frmViewDetail.cmd.value = '';
	document.frmViewDetail.DocNum.value = '';
	document.frmViewDetail.isEntry.value = IsEntry;
	document.frmViewDetail.CardCode.value = '';
	document.frmViewDetail.ViewOnly.value = 'Y';
	document.frmViewDetail.high.value = high;
	if (linkCat == 'N')
	{
		if (ObjectCode == 2 || ObjectCode == -7)
		{
			document.frmViewDetail.AddPath.value = '../';
			
			if (ObjectCode == 2)
				document.frmViewDetail.CardCode.value = Entry;
			else if (ObjectCode == -7)
			{
				document.frmViewDetail.ViewOnly.value = 'N';
				document.frmViewDetail.DocEntry.value = Entry;			
			}
				
			document.frmViewDetail.action = imgAddPath + 'addCard/crdConfDetailOpen.asp';
		}
		else if (ObjectCode == -9 || ObjectCode == 33)
		{
			document.frmViewDetail.AddPath.value = '../';
			document.frmViewDetail.ViewOnly.value = 'N';
			document.frmViewDetail.DocEntry.value = Entry;	
			document.frmViewDetail.action = imgAddPath + 'addActivity/activityConfDetail.asp';						
		}
		else if (ObjectCode == -11 || ObjectCode == 97)
		{
			document.frmViewDetail.AddPath.value = '../';
			document.frmViewDetail.ViewOnly.value = 'N';
			document.frmViewDetail.DocEntry.value = Entry;	
			document.frmViewDetail.action = imgAddPath + 'addSO/soConfDetail.asp';						
		}
		else if (ObjectCode == -8)
		{
			document.frmViewDetail.AddPath.value = '../';
			document.frmViewDetail.ViewOnly.value = 'N';
			document.frmViewDetail.DocEntry.value = Entry;	
			document.frmViewDetail.action = imgAddPath + 'addItem/itmConfDetail.asp';			
		}
		else if (ObjectCode == 4)
		{
			switch (UserType)
			{
				case 'C':
					window.location.href = 'item.asp?Item=' + Entry + '&cmd=' + itemCmd + '&pop=Y&AddPath=../';
					return;
				case 'V':
					ItemCmd = repItemCmd;
					openItemDetails(Entry);
					return;
			}
		}
		else if (ObjectCode == 24 || ObjectCode == 46 || ObjectCode == -6 || ObjectCode == 140)
		{
			document.frmViewDetail.action = imgAddPath + 'cxcRctDetailOpen.asp';
		}
		else if (ObjectCode == 13 || ObjectCode == 17 || ObjectCode == 23 || ObjectCode == 15 || ObjectCode == 16 || ObjectCode == 14 ||
				ObjectCode == 22 || ObjectCode == 20 || ObjectCode == 21 || ObjectCode == 18 || ObjectCode == 19 || ObjectCode == -4 || ObjectCode == 112)
		{
			document.frmViewDetail.action = imgAddPath + 'cxcDocDetailOpen.asp';
		}
		else if (ObjectCode == -10)
		{
			if (linkPop == 'N')
			{
				switch (UserType)
				{
					case 'V':
						document.frmViewDetail.action = 'ventas/gocxc.asp';
						break;
					case 'C':
						if (Entry != UserName)
						{
							alert(DtxtRestData);
							return;
						}
						document.frmViewDetail.action = 'cxc.asp';
						break;
				}
				document.frmViewDetail.target = '_self';
			}
			else
			{
				document.frmViewDetail.action = imgAddPath + 'cxcPrint.asp';
				document.frmViewDetail.target = '_blank';
			}
			document.frmViewDetail.c1.value = Entry;
		}
		else if (ObjectCode != -4)
		{
			alert('LtxtDynObjErr'.replace('{0}', ObjectCode));
			return false;
		}
	
		if (ObjectCode == 24 || ObjectCode == 13 || ObjectCode == 17 || ObjectCode == 23 || ObjectCode == 15 || ObjectCode == 16 || ObjectCode == 14 ||
				ObjectCode == 22 || ObjectCode == 20 || ObjectCode == 21 || ObjectCode == 18 || ObjectCode == 19 || ObjectCode == 46 || ObjectCode == -4 || ObjectCode == -6 || ObjectCode == 112 || ObjectCode == 140)
		{
			document.frmViewDetail.DocEntry.value = Entry;
		}
	
		if (ObjectCode == -4 || ObjectCode == -6 || ObjectCode == -7 || ObjectCode == -8 || ObjectCode == -9 || ObjectCode == -11)
		{
			document.frmViewDetail.DocType.value = -2;
			document.frmViewDetail.submit();
		}
		else if (ObjectCode != 4)
		{
			document.frmViewDetail.DocType.value = ObjectCode;
			document.frmViewDetail.submit();
		}
		
	}
	else
	{
		document.frmViewDetail.sourceDoc.value = ObjectCode;
		document.frmViewDetail.DocNum.value = Entry;
		document.frmViewDetail.action = 'search.asp';//!itemSmallRep ? 'search.asp' : 'viewReportPrint.asp';
		document.frmViewDetail.cmd.value = 'searchCatalog';
		document.frmViewDetail.submit();
	}
}

function Start(page) {
	switch (UserType)
	{
		case 'C':
			OpenWin = this.open(page, "searchCart", "toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no,width=482,height=" + (UserName != '-Anon-' ? 450 : 340));
			break;
		case 'V':
			OpenWin = this.open(page, "searchCart", "toolbar=no,menubar=no,location=no,resizable=no,scrollbars=yes,width=598,height=510,status=yes");
			break;
	}
	OpenWin.focus()
}

function goObjAct(actionID, objectCode, entry)
{
	document.frmGoAction.ID.value = actionID;
	document.frmGoAction.ObjectCode.value = objectCode;
	document.frmGoAction.Entry.value = entry;
	var myFunc = '';
	switch (parseInt(actionID))
	{
		case 2:
			myFunc = 'O2';
			break;
		case 3:
			myFunc = 'O3';
			break;
		case 6:
			myFunc = 'O4';
			break;
	}
	setFlowAlertVars(myFunc, (objectCode + '{S}' + entry), 'document.frmGoAction.submit();', '');
	doFlowAlert();
}

function goApproveOrder(docEntry)
{
	document.frmGoAction.ID.value = 0;
	document.frmGoAction.ObjectCode.value = 17;
	document.frmGoAction.Entry.value = docEntry;
	setFlowAlertVars('O0', docEntry, 'document.frmGoAction.submit()', '');
	doFlowAlert();
}

function goConvQuoteOrder(docEntry, series)
{
	document.frmGoAction.ID.value = 1;
	document.frmGoAction.ObjectCode.value = 17;
	document.frmGoAction.Entry.value = docEntry;
	document.frmGoAction.Series.value = series;
	setFlowAlertVars('O1', (docEntry + '{S}' + series), 'document.frmGoAction.submit()', '');
	doFlowAlert();
}

function goConvOrderInvoice(docEntry, series)
{
	document.frmGoAction.ID.value = 7;
	document.frmGoAction.ObjectCode.value = 13;
	document.frmGoAction.Entry.value = docEntry;
	document.frmGoAction.Series.value = series;
	setFlowAlertVars('O7', (docEntry + '{S}' + series), 'document.frmGoAction.submit()', '');
	doFlowAlert();
}

function goAddItem(item, qty, unit, price, locked, whsCode)
{
	document.frmGoAddItem.Item.value = item;
	if (qty != null) document.frmGoAddItem.T1.value = qty;
	document.frmGoAddItem.SaleType.value = unit;
	document.frmGoAddItem.precio.value = price;
	document.frmGoAddItem.Locked.value = locked;
	document.frmGoAddItem.WhsCode.value = whsCode;
	setFlowAlertVars('D2', (item + '{S}' + qty + '{S}' + unit + '{S}' + price + '{S}' + whsCode), 'document.frmGoAddItem.DocConf.value = typeIDs; document.frmGoAddItem.submit()', '');
	doFlowAlert();
}

function goAddWish(item)
{
	document.frmGoAddWish.Item.value = item;
	document.frmGoAddWish.submit();
}

function reloadRep()
{
	document.frmReload.btnReload.click();
}
function saveRepPdf(Excell)
{
	if (Excell == 'N')
		document.frmReload.action = 'portal/viewRepPDF.asp';
	else
		document.frmReload.action = 'portal/viewReportPDF.asp';
		
	document.frmReload.target = '_blank';
	document.frmReload.Excell.value = Excell;
	document.frmReload.submit();
	document.frmReload.target = '';
	document.frmReload.action = 'default.asp'
}


function printShowLegend(val)
{
	if (doRepLegend) document.getElementById('btnShowLegend').style.display = val ? '' : 'none';
	if (UserType == 'C') document.getElementById('tblRepButtons').style.display = val ? '' : 'none';
}

if (Refresh > 0)
{
	setTimeout("reloadRep()", Refresh*60000);
}

if (UserType == 'V') isShowRep = true;
