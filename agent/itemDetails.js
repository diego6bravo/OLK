var itemDetailsID;
var itemDetailsNumIn;
var itemDetailsSalPack;
var itemDetailsTreeType;
var itemDetailsSaleType;
var itemDetailsManPrc;
var itemDetailsCompType;
var itemDetailsCompCount = 0;
var itemDetailsFatherUnitPrice;
var itemDetailsCur;
var itemDetailsNumInSale;
var itemDetailsSalPack
var itemLockWhs;
var itemLockWhsIndex;
var itemLoadLineID = -1;
var itemLoadWhs = '';
var olkCombo;
var hideComp;
var arrItemRep;
var arrWhs;
var vDisp;
var enableItemRep;
var enableItemInvRep;
var enableItemLineDisc;
var refreshAfterAdd;

var olkComboShowComp;
var olkComboShowFatherPrice;
var olkComboAllowChangeFatherPrice;
var olkComboShowCompPrice;
var olkComboAllowChangeCompPrice;
var olkComboVirtual;

function itemDetailsFilterCart()
{
	doMyLink('cart.asp', 'string=' + itemDetailsID + '&document=B', '');
}

function doItemRepLink(rowIndex, rsIndex)
{
	var WhsCode = document.getElementById('itemDetailsWhs').value;

	$.post("Fetch/itemDetailsFetch.asp?d=" + (new Date()).toString(), { DataType: 'IRD', rowIndex: rowIndex, rsIndex: rsIndex, Item: itemDetailsID, WhsCode: WhsCode },
	   	function(data)
	   	{
	   		var arrData = data.split('{S}');
	   		var rsVars = '';
	   		for (var i = 0;i<arrData.length-1;i++)
	   		{
	   			var arrVals = arrData[i].split('{C}');
	   			
	   			var varValue;
	   			
				rsVars += '&var' + arrVals[0] + '=';
				
				switch (arrVals[2])
				{
					case 'V':
						switch (arrVals[5])
						{
							case 'DP':
								varValue = arrVals[4];
								break;
							default:
								varValue = arrVals[3];
								break;
						}
						break;
					case 'Q':
						varValue = arrVals[3];
						break;
					case 'F':
						if (ItemCmd == 'A')
						{
							switch (arrVals[3])
							{
								case '@Price':
									varValue = document.getElementById('txtItemDetailsPrice').value.replace(itemDetailsCur, '');
									break;
								case '@Quantity':
									varValue = document.getElementById('txtItemDetailsQty').value;
									break;
								case '@Unit':
									varValue = document.getElementById('txtItemDetailsSaleType').value;
									break;
								default:
									varValue = '';
									break;
							}
						}
						else
							varValue = '';
						break;
				}
				
				rsVars += varValue;
				
				if (arrVals[5] == 'DD' || arrVals[5] == 'L')
				{
					rsVars += '&var' + arrVals[0] + 'Desc=';
				}
				
	   		}

		   var wOpen;
		   var sOptions;
		
		   sOptions = 'toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes';
		   sOptions = sOptions + ',width=' + (screen.availWidth - 10).toString();
		   sOptions = sOptions + ',height=' + (screen.availHeight - 122).toString();
		   sOptions = sOptions + ',screenX=0,screenY=0,left=0,top=0';
		
		   wOpen = window.open('', 'repLinkDet', sOptions );
		   //wOpen.location = page;
		   wOpen.focus();
		   wOpen.moveTo( 0, 0 );
		   wOpen.resizeTo( screen.availWidth, screen.availHeight );
		   OpenWin = wOpen;
   
	   		doMyLink('viewReportPrint.asp', 'addPath=&pop=Y&cmd=report&rsIndex=' + rsIndex + rsVars, 'repLinkDet');
		});
}

function itemDetailsAdd()
{
	var Quantity = document.getElementById('txtItemDetailsQty').value;
	var Unit = document.getElementById('txtItemDetailsSaleType').value;
	var Price = document.getElementById('txtItemDetailsPrice').value;
	var WhsCode = document.getElementById('itemDetailsWhs').value;
	var chkAddAll = document.getElementById('txtItemDetailsAddAll') ? (document.getElementById('txtItemDetailsAddAll').checked ? 'Y' : 'N') : 'N';
	
	setFlowAlertVars('D2', (itemDetailsID + '{S}' + Quantity + '{S}' + Unit + '{S}' + clearSymbol(Price, itemDetailsCur) + '{S}' + WhsCode + '{S}' + chkAddAll), 'itemDetailsExecuteAdd();', '');
	doFlowAlert();
}

function clearSymbol(value, symbol)
{
	var retVal = value.replace(symbol, '');
	retVal = retVal.substr(1, retVal.length-1);
	return retVal;
}

function itemDetailsAddWL()
{
	$.post("Fetch/itemDetailsFetch.asp?d=" + (new Date()).toString(), { DataType: 'AWL', Item: itemDetailsID },
	   function(data)
	   {
	   		var arrData = data.split('{S}');
	   		
	   		switch (arrData[0])
	   		{
	   			case 'ok':
	   				alert(txtItemAddOK);
	   				closeItemDetails();
	   				break;
	   		}
	   });
}


function itemDetailsExecuteAdd()
{
	var txtItemDetailsQty = document.getElementById('txtItemDetailsQty');
	var txtItemDetailsSaleType = document.getElementById('txtItemDetailsSaleType');
	var txtItemDetailsDisc = document.getElementById('txtItemDetailsDisc');
	var txtItemDetailsPrice = document.getElementById('txtItemDetailsPrice');
	var itemDetailsWhs = document.getElementById('itemDetailsWhs');
	var txtItemDetailsMemo = document.getElementById('txtItemDetailsMemo');
	var chkAddAll = document.getElementById('txtItemDetailsAddAll') ? (document.getElementById('txtItemDetailsAddAll').checked ? 'Y' : 'N') : 'N';
	var selTaxCode = '';
	if (document.getElementById('TaxCode'))
	{
		selTaxCode = document.getElementById('TaxCode').value;
	}
	
	var compData = '';
	
	if (itemDetailsCompType != '')
	{
		for (var i = 0;i<itemDetailsCompCount;i++)
		{
			var itemCode = document.getElementById('compItem' + i);
			var compChildID = document.getElementById('compChildID' + i);
			var compQty = document.getElementById('compQty' + i);
			//var compTreeType = document.getElementById('compTreeType' + i);
			var compUnit = document.getElementById('compUnit' + i);
			var compWhs = document.getElementById('compWhs' + i);
			var compDiscount = document.getElementById('compDiscount' + i);
			var compPrice = document.getElementById('compPrice' + i);
			var compManPrc = document.getElementById('compManPrc' + i);
			var compCur = document.getElementById('compCur' + i);
			var compTaxCode = '';
			var compRecQty = document.getElementById('compRecQty' + i);
			var compHideComp = document.getElementById('compHideComp' + i);
			var chkComp = document.getElementById('chkComp' + i);

			if (chkComp.checked)
			{
				if (compData != '') compData += '{I}';
				compData += itemCode.value + '{S}' + compQty.value + '{S}' + compUnit.value + '{S}' + compWhs.value + '{S}' + clearSymbol(compPrice.value, compCur.value) + '{S}' + compManPrc.value + '{S}' + compTaxCode + '{S}' + compRecQty.value + '{S}' + compHideComp.value + '{S}' + compChildID.value + '{S}' + clearSymbol(compDiscount.value, '%');
			}
		}
	}
	
	$.post("Fetch/itemDetailsFetch.asp?d=" + (new Date()).toString(), { 
																			DataType: 'AI', 
																			Item: itemDetailsID, 
																			Qty: txtItemDetailsQty.value,
																			SaleType: txtItemDetailsSaleType.value,
																			Price: clearSymbol(txtItemDetailsPrice.value, itemDetailsCur),
																			DiscPrcnt: clearSymbol(txtItemDetailsDisc.value, '%'), 
																			ManPrc: (itemDetailsManPrc ? 'Y' : 'N'),
																			Whs: itemDetailsWhs.value,
																			Note: (!txtItemDetailsMemo.disabled ? txtItemDetailsMemo.value : ''), 
																			TreeType: itemDetailsTreeType, 
																			CompType: itemDetailsCompType,
																			Virtual: (olkComboVirtual ? 'Y' : 'N'),
																			HideComp: (hideComp ? 'Y' : 'N'),
																			AddAll: chkAddAll, 
																			TaxCode: selTaxCode,
																			Currency: itemDetailsCur,
																			CompData: compData
																		},
	   function(data)
	   {
	   		var arrData = data.split('{S}');
	   		
	   		switch (arrData[0])
	   		{
	   			case 'ok':
	   				//alert(txtItemAddOK);
	   				closeFlowAlert();
	   				closeItemDetails();
	   				if (!refreshAfterAdd)
	   				{
	   					loadMinRep();
	   					if (document.frmSmallSearch.string) document.frmSmallSearch.string.focus();
	   				}
	   				else
	   				{
	   					window.location.href = 'cart.asp';
	   				}
	   				break;
	   			case 'ie':
	   				alert(txtInvErrMsg.replace('{0}', itemDetailsID));
	   				break;
	   		}
	   });
	
}

function setItemDetailsBtnAddDis()
{
	if (ItemCmd != 'A') return;
	
	var txtItemDetailsBtnAdd = document.getElementById('txtItemDetailsBtnAdd');
	var txtItemDetailsAddAll = document.getElementById('txtItemDetailsAddAll');
	
	var enable = true;
	
	if (document.getElementById('ItemDetailsInvErr').style.display == '') enable = false;
	
	if (enable && itemDetailsCompType != '')
	{
		for (var i = 0;i<itemDetailsCompCount;i++)
		{
			if (document.getElementById('compInvErr' + i).style.display == '' && document.getElementById('chkComp' + i).checked)
			{
				enable = false;
				break;
			}
		}
	}
	
	txtItemDetailsBtnAdd.disabled = !enable;
	if (txtItemDetailsAddAll) txtItemDetailsAddAll.disabled = !enable;
}

function LoadItemDetailsCommRep(s)
{
	var tabID = s == 'O' ? '6' : '7';
	
	if (document.getElementById('itemDetTabs-' + tabID).innerHTML != '')
	{
		LoadItemCommRep(s);
	}
	else
	{
		GenerateItemCommRep(s);
	}
}

function LoadItemCommRep(s)
{
	showItemDetailAJAXLoader('Whole', true);
	
	$.post('Fetch/itemDetailsFetch.asp?d=' + (new Date()).toString(), { DataType: 'CR', Item: itemDetailsID, Source: s }, function(data)
	{
		var trNoData = document.getElementById('tblItemDetCommRep' + s + 'NoData');
		var tBodyData = document.getElementById('tblItemDetCommRep' + s + 'Data');
		
		if (data != 'nodata')
		{
			var arrData = data.split('{S}');
			
			var strItemComRepData = '';
			
			for (var i = 0;i<arrData.length;i++)
			{
				var arrValues = arrData[i].split('{C}');
				
				strItemComRepData += '<tr class="GeneralTbl" style="">' ;
				strItemComRepData += '	<td width="20"><a href="javascript:olkOpenObj(' + arrValues[0] + ', ' + arrValues[1] + ', ' + arrValues[2] + ');"><img border="0" src="design/0/images/' + rtl + 'felcahSelect.gif" width="15" height="13"></a></td>' ;
				strItemComRepData += '	<td class="nobr">' + arrValues[3] + '&nbsp;</td>' ;
				strItemComRepData += '	<td class="nobr">';
				
				switch (parseInt(arrValues[4]))
				{
					case 13:
						strItemComRepData += txtInv;
						break;
					case 17:
						strItemComRepData += txtOrdr;
						break;
				}
				
				strItemComRepData += '</td>' ;
				strItemComRepData += '	<td class="nobr">' + arrValues[5] + '&nbsp;</td>' ;
				strItemComRepData += '	<td class="nobr">' + arrValues[6] + '&nbsp;</td>' ;
				strItemComRepData += '	<td class="nobr">' + arrValues[7] + '&nbsp;</td>' ;
				strItemComRepData += '	<td class="nobr">' + arrValues[8] + '&nbsp;</td>' ;
				strItemComRepData += '	<td align="right" class="nobr">&nbsp;' + arrValues[9] + '</td>' ;
				strItemComRepData += '	<td align="center" class="nobr">' + arrValues[10] + '</td>' ;
				strItemComRepData += '	<td align="right" class="nobr">&nbsp;' + arrValues[11] + '</td>' ;
				strItemComRepData += '</tr>' ;
			}
			
			trNoData.style.display = 'none';
			
			$('#tblItemDetCommRep' + s + 'Data').html(strItemComRepData);
			tBodyData.style.display = '';
		}
		else
		{
			trNoData.style.display = '';
			tBodyData.style.display = 'none';
		}
		
		showItemDetailAJAXLoader('Whole', false);
	});
}

function GenerateItemCommRep(s)
{
	var tabID = s == 'O' ? '6' : '7';
	
	var strCommRep = '';
	
	strCommRep += '<table border="0" cellpadding="0" width="100%">' ;
	strCommRep += '	<tr class="GeneralTblBold2">' ;
	strCommRep += '		<td align="center" colspan="2">#</td>' ;
	strCommRep += '		<td align="center">' + txtType + '</td>' ;
	strCommRep += '		<td align="center">' + txtDate + '</td>' ;
	strCommRep += '		<td align="center">' + txtAgent + '</td>' ;
	strCommRep += '		<td align="center">' + lblItemDetailsCode + '</td>' ;
	strCommRep += '		<td align="center">' + txtClient + '</td>' ;
	strCommRep += '		<td align="center">' + lblItemDetailsQty + '</td>' ;
	strCommRep += '		<td align="center">' + lblItemDetailsUnit + '</td>' ;
	strCommRep += '		<td align="center">' + lblItemDetailsPrice + '</td>' ;
	strCommRep += '	</tr>' ;
	strCommRep += '	<tbody id="tblItemDetCommRep' + s + 'Data"></tbody>' ;
	strCommRep += '	<tr class="GeneralTbl" id="tblItemDetCommRep' + s + 'NoData">' ;
	strCommRep += '		<td colspan="10" align="center">' + txtNoData + '</td>' ;
	strCommRep += '	</tr>' ;
	strCommRep += '</table>' ;
	
	document.getElementById('itemDetTabs-' + tabID).innerHTML = strCommRep;
	
	LoadItemCommRep(s);
}

function LoadItemDetailsSalesRep()
{
	if (document.getElementById('itemDetTabs-5').innerHTML != '')
	{
		LoadItemSalesRep();
	}
	else
	{
		GenerateItemSalesRep();
	}
}

function LoadItemSalesRep()
{
	showItemDetailAJAXLoader('Whole', true);
	
	
	$.post('Fetch/itemDetailsFetch.asp?d=' + (new Date()).toString(), { DataType: 'SR', Item: itemDetailsID }, function(data)
	{
		var trNoData = document.getElementById('tblItemDetSalesRepNoData');
		var tBodyData = document.getElementById('tblItemDetSalesRepData');
		
		if (data != 'nodata')
		{
			var arrData = data.split('{S}');
			
			var strSalesRepData = '';
			
			for (var i = 0;i<arrData.length;i++)
			{
				var arrValues = arrData[i].split('{C}');
				
				strSalesRepData += '	<tr class="GeneralTbl">' ;
				strSalesRepData += '		<td>' ;
				strSalesRepData += '		<a href="javascript:olkOpenObj(13,' + arrValues[0] + ', ' + arrValues[1] + ')"><img border="0" src="design/0/images/' + rtl + 'felcahSelect.gif" width="15" height="13"></a></td>' ;
				strSalesRepData += '		<td>' + arrValues[2] + '</td>' ;
				strSalesRepData += '		<td>' + arrValues[3] + '&nbsp;</td>' ;
				strSalesRepData += '		<td>' + arrValues[4] + '&nbsp;</td>' ;
				strSalesRepData += '		<td>' + arrValues[5] + '&nbsp;</td>' ;
				strSalesRepData += '		<td align="right">' + arrValues[6] + '&nbsp;</td>' ;
				strSalesRepData += '		<td>' + arrValues[7] + '&nbsp;</td>' ;
				strSalesRepData += '		<td align="right"><nobr>' + arrValues[8] + '</nobr></td>' ;
				strSalesRepData += '	</tr>' ;
			}
			
			trNoData.style.display = 'none';
			
			$('#tblItemDetSalesRepData').html(strSalesRepData);
			tBodyData.style.display = '';
		}
		else
		{
			trNoData.style.display = '';
			tBodyData.style.display = 'none';
		}
		
		showItemDetailAJAXLoader('Whole', false);
	});
}

function GenerateItemSalesRep()
{
	var strSalesRep;
	strSalesRep = '<table border="0" cellpadding="0" width="100%">' ;
	strSalesRep += '	<tr class="GeneralTblBold2">' ;
	strSalesRep += '		<td align="center" colspan="2">#</td>' ;
	strSalesRep += '		<td align="center">' + txtDate + '</td>' ;
	strSalesRep += '		<td align="center">' + lblItemDetailsCode + '</td>' ;
	strSalesRep += '		<td align="center">' + txtClient + '</td>' ;
	strSalesRep += '		<td align="center">' + lblItemDetailsQty + '</td>' ;
	strSalesRep += '		<td align="center">' + txtSalMet + '</td>' ;
	strSalesRep += '		<td align="center">' + lblItemDetailsPrice + '</td>' ;
	strSalesRep += '	</tr>' ;
	strSalesRep += '	<tbody id="tblItemDetSalesRepData" style="display: none;"></tbody>' ;
	strSalesRep += '	<tr class="GeneralTbl" id="tblItemDetSalesRepNoData">' ;
	strSalesRep += '		<td colspan="8">' ;
	strSalesRep += '		<p align="center">' + txtNoData + '</td>' ;
	strSalesRep += '	</tr>' ;
	strSalesRep += '</table>' ;
	
	$('#itemDetTabs-5').html(strSalesRep);
	
	LoadItemSalesRep();

}

function LoadItemDetailsBdgDetailsRep()
{
	if (document.getElementById('itemDetTabs-3').innerHTML != '')
	{
		LoadItemBdgDetailsRep();
	}
	else
	{
		GenerateItemBdgDetailsRep();
	}
}

function LoadItemBdgDetailsRep()
{
	showItemDetailAJAXLoader('Whole', true);
	
	var whsCode = document.getElementById('itemDetailsWhs').value;
	
	if (!itemLockWhs)
	{
		whsCode = '';
		for (var i = 0;i<arrWhs.length;i++) {
			if (i > 0) whsCode += ', ';
			whsCode += arrWhs[i].split('{C}')[0];
		}
	}
	
	$.post('Fetch/itemDetailsFetch.asp?d=' + (new Date()).toString(), { DataType: 'WR', Item: itemDetailsID, WhsCode: whsCode }, function(data)
	{
		$('#itemDetInvRepUn2Desc').text(itemDetailsNumInSale);
		$('#itemDetInvRepUn3Desc').text(itemDetailsSalPack);
		
		if (!itemLockWhs)
		{
			for (var i = 0;i<arrWhs.length;i++)
			{
				$('#itemDetInvRepUn2Desc_' + i).text(itemDetailsNumInSale);
				$('#itemDetInvRepUn3Desc_' + i).text(itemDetailsSalPack);
			}
		}
		else
		{
			$('#itemDetInvRepUn2Desc_' + itemLockWhsIndex).text(itemDetailsNumInSale);
			$('#itemDetInvRepUn3Desc_' + itemLockWhsIndex).text(itemDetailsSalPack);
		}
		
		var arrData = data.split('{S}');
		
		var arrAll = arrData[0].split('{C}');
		
		$('#itemDetInvRepOnHand').text(arrAll[0]);
		$('#itemDetInvRepDispSAP').text(arrAll[1]);
		$('#itemDetInvRepInvOLKDisp').text(arrAll[2]);
		$('#itemDetInvRepOnHandUnVentSAP').text(arrAll[3]);
		$('#itemDetInvRepDispUnVentSAP').text(arrAll[4]);
		$('#itemDetInvRepInvOLKUnVentDisp').text(arrAll[5]);
		$('#itemDetInvRepOnHandUnEmbSAP').text(arrAll[6]);
		$('#itemDetInvRepDispUnEmbSAP').text(arrAll[7]);
		$('#itemDetInvRepInvOLKUnEmbDisp').text(arrAll[8]);
		
		var arrRep = arrData[1].split('{W}');
		
		if (!itemLockWhs)
		{
			for (var i = 0;i<arrWhs.length;i++)
			{
				LoadItemBdgDetailsRepData(arrRep[i], i);
			}
		}
		else
		{
			LoadItemBdgDetailsRepData(arrRep[0], itemLockWhsIndex);
		}
		
		showItemDetailAJAXLoader('Whole', false);
	});
}

function LoadItemBdgDetailsRepData(data, i)
{
	var arrData = data.split('{C}');

	$('#itemDetInvRepInvBDGWhs_' + i).text(arrData[0]);
	$('#itemDetInvRepInvBDGDisp_' + i).text(arrData[1]);
	$('#itemDetInvRepInvOLKBDGDisp_' + i).text(arrData[2]);
	$('#itemDetInvRepInvUnVentBDGWhs_' + i).text(arrData[3]);
	$('#itemDetInvRepInvBDGUnVentDisp_' + i).text(arrData[4]);
	$('#itemDetInvRepInvOLKBDGUnVentDisp_' + i).text(arrData[5]);
	$('#itemDetInvRepInvUnEmbBDGWhs_' + i).text(arrData[6]);
	$('#itemDetInvRepInvBDGUnEmbDisp_' + i).text(arrData[7]);
	$('#itemDetInvRepInvOLKBDGUnEmbDisp_' + i).text(arrData[8]);
}

function GenerateItemBdgDetailsRep()
{
	var strItemInvRep;
	strItemInvRep = '<table border="0" cellpadding="0" width="100%">' ;
	strItemInvRep += '<tr>' ;
	strItemInvRep += '<td>' ;
	strItemInvRep += '<table border="0" cellpadding="0" width="100%" cellspacing="1">' ;
	strItemInvRep += '<tr>' ;
	strItemInvRep += '<td valign="top">' ;
	strItemInvRep += '<table border="0" cellpadding="0" width="100%">' ;
	strItemInvRep += '<tr class="GeneralTblBold2" style="text-align: center;">' ;
	strItemInvRep += '<td colspan="2">' + txtOnHand + '</td>' ;
	strItemInvRep += '<td colspan="2" style="width: 199px;">' + txtAVL + '</td>' ;
	strItemInvRep += '</tr>' ;
	strItemInvRep += '<tr class="GeneralTbl">' ;
	strItemInvRep += '<td class="GeneralTblBold2">&nbsp;</td>' ;
	strItemInvRep += '<td class="GeneralTblBold2">' ;
	strItemInvRep += txtSAP + '</td>' ;
	strItemInvRep += '<td style="width: 100px;" class="GeneralTblBold2">' + txtSAP + '</td>' ;
	strItemInvRep += '<td style="width: 99px;" class="GeneralTblBold2">' + txtOLK + '</td>' ;
	strItemInvRep += '</tr>' ;
	strItemInvRep += '<tr class="GeneralTbl">' ;
	strItemInvRep += '<td class="GeneralTblBold2">' + lblItemDetailsUnit + '</td>' ;
	strItemInvRep += '<td><span id="itemDetInvRepOnHand"></span>&nbsp;</td>' ;
	strItemInvRep += '<td style="width: 100px;"><span id="itemDetInvRepDispSAP"></span>&nbsp;</td>' ;
	strItemInvRep += '<td style="width: 99px;"><span id="itemDetInvRepInvOLKDisp"></span>&nbsp;</td>' ;
	strItemInvRep += '</tr>' ;
	strItemInvRep += '<tr class="GeneralTbl">' ;
	strItemInvRep += '<td class="GeneralTblBold2"><span id="itemDetInvRepUn2Desc"></span></td>' ;
	strItemInvRep += '<td><span id="itemDetInvRepOnHandUnVentSAP"></span></td>' ;
	strItemInvRep += '<td style="width: 100px;"><span id="itemDetInvRepDispUnVentSAP"></span></td>' ;
	strItemInvRep += '<td style="width: 99px;"><span id="itemDetInvRepInvOLKUnVentDisp"></span></td>' ;
	strItemInvRep += '</tr>' ;
	strItemInvRep += '<tr class="GeneralTbl">' ;
	strItemInvRep += '<td class="GeneralTblBold2"><span id="itemDetInvRepUn3Desc"></span></td>' ;
	strItemInvRep += '<td><span id="itemDetInvRepOnHandUnEmbSAP"></span></td>' ;
	strItemInvRep += '<td style="width: 100px;"><span id="itemDetInvRepDispUnEmbSAP"></span></td>' ;
	strItemInvRep += '<td style="width: 99px;"><span id="itemDetInvRepInvOLKUnEmbDisp"></span></td>' ;
	strItemInvRep += '</tr>' ;
	strItemInvRep += '</table>' ;
	strItemInvRep += '</td>' ;
	strItemInvRep += '</tr>' ;
	strItemInvRep += '</table>' ;
	strItemInvRep += '</td>' ;
	strItemInvRep += '</tr>' ;
	strItemInvRep += '<tr class="GeneralTbl">' ;
	strItemInvRep += '<td><hr size="1" /></td>' ;
	strItemInvRep += '</tr>' ;
	strItemInvRep += '<tr class="GeneralTbl">' ;
	strItemInvRep += '<td>' ;
	strItemInvRep += '<table border="0" cellpadding="0" width="100%">' ;
	strItemInvRep += '<tr>' ;
	
	//strItemInvRep += '<% while not rw.eof varx = varx + 1 cmd("@WhsCode") = rw("WhsCode") set rs = cmd.execute() %>' ;
	
	var itemWhs = document.getElementById('itemDetailsWhs').value;
	
	var iCol = 0;
	for (var i = 0;i<arrWhs.length;i++)
	{
		var arrWhsData = arrWhs[i].split('{C}');
		
		if (!itemLockWhs || itemLockWhs && arrWhsData[0] == itemWhs)
		{
			strItemInvRep += '<td>' ;
			strItemInvRep += '<div align="center">' ;
			strItemInvRep += '<table border="0" cellpadding="0" width="272" cellspacing="0">' ;
			strItemInvRep += '<tr class="GeneralTlt">' ;
			strItemInvRep += '<td>' ;
			strItemInvRep += '<table border="0" cellpadding="0" width="272" cellspacing="1">' ;
			strItemInvRep += '<tr class="GeneralTbl">' ;
			strItemInvRep += '<td>' ;
			strItemInvRep += '<table border="0" cellpadding="0" width="100%" cellspacing="1">' ;
			strItemInvRep += '<tr class="GeneralTblBold2">' ;
			strItemInvRep += '<td colspan="4">' + arrWhsData[1] + '&nbsp;</td>' ;
			strItemInvRep += '</tr>' ;
			strItemInvRep += '<tr class="GeneralTblBold2">' ;
			strItemInvRep += '<td align="center" style="width: 50%;" colspan="2">' ;
			strItemInvRep += txtOnHand + '</td>' ;
			strItemInvRep += '<td align="center" style="width: 50%;" colspan="2">' ;
			strItemInvRep += txtAVL + '</td>' ;
			strItemInvRep += '</tr>' ;
			strItemInvRep += '<tr class="GeneralTblBold2">' ;
			strItemInvRep += '<td style="width: 25%;">&nbsp;</td>' ;
			strItemInvRep += '<td style="width: 25%;">' + txtWHS + '</td>' ;
			strItemInvRep += '<td style="width: 25%;">' + txtSAP + '</td>' ;
			strItemInvRep += '<td style="width: 25%;">' + txtOLK + '</td>' ;
			strItemInvRep += '</tr>' ;
			strItemInvRep += '<tr class="GeneralTbl">' ;
			strItemInvRep += '<td style="width: 25%;" class="GeneralTblBold2">' + lblItemDetailsUnit + '</td>' ;
			strItemInvRep += '<td style="width: 25%;"><span id="itemDetInvRepInvBDGWhs_' + i + '"></span>&nbsp;</td>' ;
			strItemInvRep += '<td style="width: 25%;"><span id="itemDetInvRepInvBDGDisp_' + i + '"></span>&nbsp;</td>' ;
			strItemInvRep += '<td style="width: 25%;"><span id="itemDetInvRepInvOLKBDGDisp_' + i + '"></span>&nbsp;</td>' ;
			strItemInvRep += '</tr>' ;
			strItemInvRep += '<tr class="GeneralTbl">' ;
			strItemInvRep += '<td style="width: 25%;" class="GeneralTblBold2"><span id="itemDetInvRepUn2Desc_' + i + '"></span></td>' ;
			strItemInvRep += '<td style="width: 25%;"><span id="itemDetInvRepInvUnVentBDGWhs_' + i + '"></span>&nbsp;</td>' ;
			strItemInvRep += '<td style="width: 25%;"><span id="itemDetInvRepInvBDGUnVentDisp_' + i + '"></span></td>' ;
			strItemInvRep += '<td style="width: 25%;"><span id="itemDetInvRepInvOLKBDGUnVentDisp_' + i + '"></span></td>' ;
			strItemInvRep += '</tr>' ;
			strItemInvRep += '<tr class="GeneralTbl">' ;
			strItemInvRep += '<td style="width: 25%;" class="GeneralTblBold2"><span id="itemDetInvRepUn3Desc_' + i + '"></span></td>' ;
			strItemInvRep += '<td style="width: 25%;"><span id="itemDetInvRepInvUnEmbBDGWhs_' + i + '"></span></td>' ;
			strItemInvRep += '<td style="width: 25%;"><span id="itemDetInvRepInvBDGUnEmbDisp_' + i + '"></span></td>' ;
			strItemInvRep += '<td style="width: 25%;"><span id="itemDetInvRepInvOLKBDGUnEmbDisp_' + i + '"></span></td>' ;
			strItemInvRep += '</tr>' ;
			strItemInvRep += '</table>' ;
			strItemInvRep += '</td>' ;
			strItemInvRep += '</tr>' ;
			strItemInvRep += '</table>' ;
			strItemInvRep += '</td>' ;
			strItemInvRep += '</tr>' ;
			strItemInvRep += '</table>' ;
			strItemInvRep += '</div>' ;
			strItemInvRep += '</td>' ;
			
			if (iCol++ == 2)
			{
				strItemInvRep += '</tr><tr>' ;
				iCol = 0;
			}
		}
	}
	
	strItemInvRep += '</tr>' ;
	strItemInvRep += '</table>' ;
	strItemInvRep += '</td>' ;
	strItemInvRep += '</tr>' ;
	strItemInvRep += '</table>' ;
	
	$('#itemDetTabs-3').html(strItemInvRep);
	
	LoadItemBdgDetailsRep();
}

function itemDetailsOpenPic()
{
	Pic('thumb/default.asp?item=' + itemDetailsID + '&amp;pop=Y&amp;AddPath=',529,510,'yes','yes');
}

function LoadItemDetailsData(item)
{
	itemDetailsID = item;
	
	showItemDetailAJAXLoader('Top', true);
	showItemDetailAJAXLoader('IR', true);	
	if (enableItemInvRep) showItemDetailAJAXLoader('Inv', true);
	else showItemDetailAJAXLoader('Inv', false);
	showItemDetailAJAXLoader('Bottom', true);

	$.post("Fetch/itemDetailsFetch.asp?d=" + (new Date()).toString(), { DataType: 'D', Item: item, cmd: ItemCmd, LineNum: itemLoadLineID },
   function(data)
   {
		itemDetailsManPrc = false;
		
   		var arrData = data.split('{S}');
		jQuery('#dvItemDetails').dialog("option", "title", item + ' - ' + arrData[0] );
   		document.getElementById('itemDetailsText').innerHTML = arrData[1];
   		
		var itemImg = arrData[2];

		var itemDetailsImg = document.getElementById('itemDetailsImg');
		
		if (itemImg != '')
		{
			itemDetailsImg.src = 'pic.aspx?filename=' + itemImg + '&dbName=' + dbName;
		}
		else
		{
			itemDetailsImg.src = 'pic.aspx?filename=n_a.gif&dbName=' + dbName;
		}
		
		itemDetailsTreeType = arrData[3];
		itemDetailsSaleType = arrData[13];
		
		var UnitMsr = arrData[4];
		itemDetailsNumIn = arrData[5];
		var PackMsr = arrData[6];
		itemDetailsSalPack = arrData[7];
		var WhsCode = itemLoadWhs == '' ? arrData[8] : itemLoadWhs;
		var VerfyOnWL = arrData[9] == 'Y';
		var VerfyOnCart = arrData[10] == 'Y';
		var childSum = parseFloat(getNumeric(arrData[11]));;
		var VerfyChild = childSum > 0;
		var HasVolDisc = arrData[12] == 'Y';
		var VerfyCartItemFilter = EnableCartSum && VerfyOnCart;
		var SaleType = arrData[13];
		var CartQuantity = arrData[14];
		var Price = arrData[15];
		var DiscPrcnt = arrData[16];
		var chkInv = arrData[17] == 'Y';
		olkCombo = arrData[18] == 'Y';
		hideComp = arrData[19] == 'Y';
		olkComboShowComp = arrData[20] == 'Y';
		olkComboShowFatherPrice = arrData[21] == 'Y';
		olkComboAllowChangeFatherPrice = arrData[22] == 'Y';
		olkComboShowCompPrice = arrData[23] == 'Y';
		olkComboAllowChangeCompPrice = arrData[24] == 'Y';
		olkComboVirtual = arrData[25] == 'Y';
		itemDetailsCur = arrData[26];
		var selTaxCode = arrData[27];
		
		var olkCartQty = arrData[arrData.length-9];
		var itemRecData = arrData[arrData.length-8];
		var virtualPrice = arrData[arrData.length-7];
		var virtualTotal = arrData[arrData.length-6];
		var virtualChkInv = arrData[arrData.length-5];
		var virtualDisc = arrData[arrData.length-4];
		var Quantity = arrData[arrData.length-3];
		var DisableControls = arrData[arrData.length-2] == 'Y';
		var LineTotal = arrData[arrData.length-1];
		
		itemDetailsNumInSale = UnitMsr + ' (' + itemDetailsNumIn + ')';
		itemDetailsSalPack = PackMsr + ' (' + itemDetailsSalPack + ')';
		
		document.getElementById('itemDetailsMsg').style.display = 'none';
		
		switch (ItemCmd)
		{
			case 'A':
		
				if (olkCombo && olkComboVirtual)
				{
					Price = virtualPrice;
					LineTotal = virtualTotal;
					chkInv = virtualChkInv;
					DiscPrcnt = virtualDisc;
					
					document.getElementById('txtItemDetailsDisc').disabled = true;
					document.getElementById('txtItemDetailsPrice').disabled = true;
					document.getElementById('itemDetailsWhs').disabled = true;
					
					itemDetailsFatherUnitPrice = virtualTotal;
				}
				else
				{
					document.getElementById('txtItemDetailsDisc').disabled = false;
					document.getElementById('txtItemDetailsPrice').disabled = false;
					document.getElementById('itemDetailsWhs').disabled = false;
				}
				
				document.getElementById('ItemDetailsInvErr').style.display = chkInv ? 'none' : '';
				
				if (enableItemInvRep) 
				{
					document.getElementById('txtItemDetailsInvUnit2').innerText = UnitMsr + ' (' + itemDetailsNumIn + ')';
					document.getElementById('txtItemDetailsInvUnit3').innerText = PackMsr + ' (' + itemDetailsSalPack + ')';
				}

				document.getElementById('optItemDetails2').innerHTML = itemDetailsNumInSale ;
				document.getElementById('optItemDetails3').innerHTML = itemDetailsSalPack;
			
	
				document.getElementById('itemDetailsWhs').value = WhsCode;
				if (itemLockWhs)
				{
					var strWhsName = '';
					
					for (var i = 0;i<arrWhs.length;i++) 
					{ 
						var arrWhsData = arrWhs[i].split('{C}'); 
						if (arrWhsData[0] == WhsCode) 
						{ 
							strWhsName = arrWhsData[1]; 
							itemLockWhsIndex = i;
							break; 
						} 
					}
					
					document.getElementById('itemWhsDesc').innerText = strWhsName;
				}
				
				document.getElementById('txtItemDetailsQty').value = Quantity;
				document.getElementById('txtItemDetailsQty').disabled = false;
				document.getElementById('txtItemDetailsSaleType').value = SaleType;
				document.getElementById('txtItemDetailsSaleType').disabled = DisableControls;
				if (enableItemLineDisc) 
				{
					document.getElementById('txtItemDetailsDisc').value = '% ' + DiscPrcnt;
					document.getElementById('txtItemDetailsDisc').disabled = DisableControls;
				}
				document.getElementById('txtItemDetailsPrice').value = itemDetailsCur + ' ' + Price;
				document.getElementById('txtItemDetailsPrice').disabled = DisableControls;
				document.getElementById('txtItemDetailsTotal').value = itemDetailsCur + ' ' + LineTotal;
				document.getElementById('txtItemDetailsTotal').style.display = (!olkCombo || olkCombo && olkComboShowFatherPrice) ? '' : 'none';
				document.getElementById('txtItemDetailsPrice').style.display = document.getElementById('txtItemDetailsTotal').style.display;
				document.getElementById('spItemDetailsTotal').style.display = document.getElementById('txtItemDetailsTotal').style.display;
				document.getElementById('spItemDetailsPrice').style.display = document.getElementById('txtItemDetailsTotal').style.display;
				
				
				if (CartQuantity != '')
				{
					document.getElementById('ttlCartQty').style.display = '';
					document.getElementById('tdCartQty').style.display = '';
					document.getElementById('txtItemDetailsCartQty').value = CartQuantity;
				}
				else
				{
					document.getElementById('ttlCartQty').style.display = 'none';
					document.getElementById('tdCartQty').style.display = 'none';
					document.getElementById('txtItemDetailsCartQty').value = 0;
				}
				for (var i = 1;i<=3;i++) document.getElementById('tblItemDetailsControlsAI' + i).style.display = '';
				document.getElementById('tblItemDetailsControlsWL').style.display = 'none';
				
				var txtItemDetailsAddAll = document.getElementById('txtItemDetailsAddAll');
				if (txtItemDetailsAddAll) txtItemDetailsAddAll.checked = false;
				
				if (document.getElementById('TaxCode'))
					if (selTaxCode != '')
					{
						document.getElementById('TaxCode').value = selTaxCode;
					}
				
				break;
			case 'W':
				for (var i = 1;i<=3;i++) document.getElementById('tblItemDetailsControlsAI' + i).style.display = 'none';
				document.getElementById('tblItemDetailsControlsWL').style.display = '';
				document.getElementById('itemDetailsWhs').value = WhsCode;
				break;
			default:
				document.getElementById('itemDetailsWhs').value = WhsCode;
				break;
		}
		
		var itemDetailsMsgRow = document.getElementById('itemDetailsMsgRow');
		if (HasVolDisc) itemDetailsMsgRow.style.backgroundColor = '#CCFF99';
		else if (VerfyOnCart || VerfyOnWL || VerfyChild) itemDetailsMsgRow.style.backgroundColor = '#FFD2A6';
		else itemDetailsMsgRow.style.backgroundColor = '';
		
		if (HasVolDisc)
		{
			document.getElementById('tblVolDisc').style.display = '';
			
			if (EnableCartSum) document.getElementById('tdItemDetailsCartFilterLink').style.display = 'none';
			document.getElementById('lblSalesHideCompTxt').style.display = 'none';
		}
		else if (VerfyOnCart || VerfyOnWL)
		{
			document.getElementById('tblItemDetailsMsg').style.display = '';
			if (EnableCartSum) document.getElementById('tdItemDetailsCartFilterLink').style.display = 'none';
			
			if (VerfyOnCart)
			{
				var newTxtItmInCart = txtItmInCart.replace('{0}', olkCartQty);
				switch (parseInt(itemDetailsSaleType))
				{
					case 1:
						newTxtItmInCart = newTxtItmInCart.replace('{1}', lblItemDetailsUnit);
						break;
					case 2:
						newTxtItmInCart = newTxtItmInCart.replace('{1}', UnitMsr);
						break;
					case 3:
						newTxtItmInCart = newTxtItmInCart.replace('{1}', PackMsr);
						break;
				}
				document.getElementById('itemDetailsMsg').innerHTML = newTxtItmInCart;
				document.getElementById('itemDetailsMsg').style.display = '';
				if (VerfyCartItemFilter)
				{
					if (EnableCartSum) document.getElementById('tdItemDetailsCartFilterLink').style.display = '';
				}
			}
			else if (VerfyOnWL)
			{
				document.getElementById('itemDetailsMsg').innerHTML = txtValOnLst;
				document.getElementById('itemDetailsMsg').style.display = '';
			}
			
			document.getElementById('tblVolDisc').style.display = 'none';
			document.getElementById('lblSalesHideCompTxt').style.display = 'none';
		}
		else if (VerfyChild)
		{
			document.getElementById('lblSalesHideCompTxt').style.display = '';
			document.getElementById('tblVolDisc').style.display = 'none';
			if (EnableCartSum) document.getElementById('tdItemDetailsCartFilterLink').style.display = 'none';
		}
		else
		{
			document.getElementById('tblVolDisc').style.display = 'none';
			if (EnableCartSum) document.getElementById('tdItemDetailsCartFilterLink').style.display = 'none';
			document.getElementById('lblSalesHideCompTxt').style.display = 'none';
		}
		
		if (itemRecData != '')
		{
			showItemDetailAJAXLoader('IR', false);	
			showItemDetailAJAXLoader('Inv', false);	
			if (!olkCombo || olkCombo && olkComboShowComp)
			{
				document.getElementById('liComponents').style.display = '';
				$('#itemDetTabs').tabs("select", (ItemCmd == 'A' ? 1 : 0) );
			}
			else
			{
				document.getElementById('liComponents').style.display = 'none';
				$('#itemDetTabs').tabs("select", 0);
			}
			
			if (document.getElementById('itemDetTabs-2').innerHTML != '')
			{
				LoadItemComponents(itemRecData);
			}
			else
			{
				GenerateItemComponents(itemRecData);
			}
		}
		else
		{
			document.getElementById('liComponents').style.display = 'none';
			$('#itemDetTabs').tabs("select", 0);
			setItemDetailsBtnAddDis();
		}
		
		if (ItemCmd == 'A')
		{
			switch (itemDetailsCompType)
			{
				case 'SaleTree':
					document.getElementById('txtItemDetailsDisc').readonly = childSum > 0;
					document.getElementById('txtItemDetailsSaleType').disabled = childSum > 0;
					document.getElementById('txtItemDetailsPrice').readonly = childSum > 0;
					document.getElementById('itemDetailsWhs').disabled = childSum > 0;
					document.getElementById('txtItemDetailsDisc').className = 'input' + (childSum > 0 ? 'Des' : '');
					document.getElementById('txtItemDetailsSaleType').className = 'input' + (childSum > 0 ? 'Des' : '');
					document.getElementById('txtItemDetailsPrice').className = 'input' + (childSum > 0 ? 'Des' : '');
					document.getElementById('itemDetailsWhs').className = 'input' + (childSum > 0 ? 'Des' : '');
					break;
				case 'OLKCombo':
					document.getElementById('txtItemDetailsDisc').readonly = !olkComboVirtual || olkComboVirtual && !olkComboAllowChangeFatherPrice;
					document.getElementById('txtItemDetailsSaleType').disabled = false;
					document.getElementById('txtItemDetailsPrice').readonly = !olkComboVirtual || olkComboVirtual && !olkComboAllowChangeFatherPrice;
					document.getElementById('itemDetailsWhs').disabled = olkComboVirtual
					document.getElementById('txtItemDetailsDisc').className = 'input' + (document.getElementById('txtItemDetailsDisc').readonly ? 'Des' : '');
					document.getElementById('txtItemDetailsSaleType').className = 'input';
					document.getElementById('txtItemDetailsPrice').className = 'input' + (document.getElementById('txtItemDetailsPrice').readonly ? 'Des' : '');
					document.getElementById('itemDetailsWhs').className = 'input' + (olkComboVirtual ? 'Des' : '');
					break;
				default:
					document.getElementById('txtItemDetailsDisc').readonly = false;
					document.getElementById('txtItemDetailsSaleType').disabled = false;
					document.getElementById('txtItemDetailsPrice').readonly = false;
					document.getElementById('itemDetailsWhs').disabled = false;
					document.getElementById('txtItemDetailsDisc').className = 'input';
					document.getElementById('txtItemDetailsSaleType').className = 'input';
					document.getElementById('txtItemDetailsPrice').className = 'input';
					document.getElementById('itemDetailsWhs').className = 'input';
					break;
			}
		}
		
		if (enableItemInvRep)
		{
			LoadItemInvRep();
		}
		
		if (enableItemRep)
		{
			LoadItemRep();
		}
		

		showItemDetailAJAXLoader('Top', false);
		showItemDetailAJAXLoader('Bottom', false);
   });
   
}

function LoadItemDetailsBP()
{
	showItemDetailAJAXLoader('Whole', true);
	$.post("Fetch/itemDetailsFetch.asp?d=" + (new Date()).toString(), { DataType: 'BP', Item: itemDetailsID },
   function(data)
   {
   		var arrData = data.split('|');
   		if (arrData[0] == 'ok')
   		{
	   		document.getElementById('txtItemDetailsBestPrice').innerHTML = arrData[1];
	   		document.getElementById('txtItemDetailsBestQty').innerHTML = arrData[2];
	   	}
	   	else
	   	{
	   		document.getElementById('txtItemDetailsBestPrice').innerHTML = '';
	   		document.getElementById('txtItemDetailsBestQty').innerHTML = '';
	   	}
	   	
	   	$('#tbCmbData').remove();
	   	if (arrData.length >= 3)
	   	if (arrData[3] != '')
	   	{
	   		var arrLines = arrData[3].split('{S}');
	   		var strBP = '';
	   		for (var i = 0;i<arrLines.length;i++)
	   		{	
	   			var arrLine = arrLines[i].split('{C}');
	   			strBP += '<tr class="GeneralTbl">' +  
							'<td>' +  
							'<p align="center">' +  
							'<a href="javascript:olkOpenObj(13,' + arrLine[0] + ', ' + (parseInt(arrLine[2])-1) + ')"><img border="0" src="design/0/images/' + rtl + 'felcahSelect.gif" width="15" height="13"></a></td>' +  
							'<td>' + arrLine[1] + '&nbsp;(' + arrLine[2] + ')</td>' +  
							'<td>' + arrLine[4] + '&nbsp;</td>' +  
							'<td align="right">&nbsp;' + arrLine[5] + '</td>' +  
							'<td align="right"><nobr>' + arrLine[3] + '</nobr></td>' +  
							'</tr>';
	   		}
			$('#tblItemDetailsBestPrice').html(strBP);
	   	}
	   	
		showItemDetailAJAXLoader('Whole', false);
   });

}

function doCheckComp(i)
{
	var chkCompImg = document.getElementById('chkCompImg' + i);
	var chkComp = document.getElementById('chkComp' + i);
	
	var checked = !chkComp.checked;
	
	chkCompImg.src = 'images/checkbox_' + (checked ? 'on' : 'off') + '.jpg';
	chkComp.checked = checked;
	
	setItemDetailsBtnAddDis();
	
	if (olkCombo && olkComboVirtual)
	{
		SumChildTotal();
	}
}
function LoadItemComponents(itemRecData)
{
	if (itemRecData == '') return;
	
	
	document.getElementById('spLblItmDetPrc').style.display = (!olkCombo || olkCombo && olkComboShowCompPrice) ? '' : 'none';
	document.getElementById('spLblItmDetTot').style.display = document.getElementById('spLblItmDetPrc').style.display;
	
	var str = '';
	
	var arrData = itemRecData.split('{C}');
	itemDetailsCompType = arrData[0];

	var arrComp = arrData[1].split('{R}');
	
	itemDetailsCompCount = arrComp.length;

	for (var i = 0;i<arrComp.length;i++)
	{
		var compData = arrComp[i].split('{O}');
		
		var ItemCode = compData[0];
		var ItemName = compData[1];
		var Quantity = compData[2];
		var Price = compData[3];
		var PicturName = compData[4];
		var DocEntry = compData[5];
		var Currency = compData[6];
		var Checked = compData[7] == 'Y';
		var Locked = compData[8] == 'Y';
		var WhsCode = compData[9];
		var Comments = compData[10];
		var compHideComp = compData[11] == 'Y';
		var SaleTypeDesc = compData[12];
		var LineTotal = compData[13];
		var DocCur = compData[14];
		var DiscType = compData[15];
		var Discount = compData[16];
		var LockPrice = compData[17] == 'Y' || (olkCombo && !olkComboAllowChangeCompPrice);
		var NumInSale = compData[18];
		var SalUnitMsr = compData[19];
		var SalPackUn = compData[20];
		var SalPackMsr = compData[21];
		var chkInv = compData[22] == 'Y';
		var compTreeType = compData[23];
		var chkVolDisc = compData[24] == 'Y';
		var RecQty = compData[25];
		var LockQty = compData[26] == 'Y';
		var childID = olkCombo ? compData[27] : '';
	
		var compSaleType = itemDetailsCompType == 'SaleTree' ? 1 : itemDetailsSaleType;
		
		
		document.getElementById('itemRecWhs').style.display = itemDetailsCompType == 'ItemRec' ? '' : 'none';
		
		str += '<tr>';
		str += '<td align="middle" class="CanastaTbl" style="width: 20px" rowspan="2"><img src="images/item_template.gif"></td>' ;
		str += '<td colspan="9" class="CanastaTbl"><table border="0" cellPadding="0" cellSpacing="0" width="100%">';
		str += '<tr class="CanastaTbl">';
		str += '<td>';
		str += '<input type="hidden" id="compItem' + i + '" value="' + ItemCode.replace('"', '""') + '"><input type="hidden" name="compChildID' + i + '" value="' + childID + '">';
		str += '<a class="LinkTop" href="#">' + ItemName + '</a></td>';
		
		if (GetShowRef)
		{
			str += '<td width="1"><nobr>';
			str += '<a class="LinkTop" href="#">' + ItemCode + '</a></nobr></td>';
		}
		
		str += '</tr>';
		str += '</table></td></tr>';
		
		str += '<tr>' ;
		str += '<td style="PADDING-LEFT: 2px; PADDING-RIGHT: 2px;">' ;
		str += '<a href="#">' ;
		str += '<img align="' + (rtl == '' ? 'left' : 'right') + '" border="0" src="pic.aspx?filename=' + (PicturName != '' ? PicturName : 'n_a.gif') + '&amp;dbName=' + dbName + '&amp;MaxSize=40" /></a><span class="CanastaTblResaltada" style="BACKGROUND-COLOR: #ffffff; FONT-WEIGHT: normal">' + Comments + '</span></td>' ;
		str += '<td style="text-align: center; padding-top: 4px; width: 20px;">' ;
		
		var loadChecked = Checked || Locked;
		
		str += '<img src="images/checkbox_' + (Locked || ItemCmd != 'A' ? 'dis_' : '') + (loadChecked ? 'on' : 'off') + '.jpg" id="chkCompImg' + i + '" border="0" ' + (!Locked ? 'onclick="javascript:doCheckComp(' + i +');"' : '') + '>';
		
		str += '<input id="chkComp' + i + '" ' + (loadChecked ? 'checked' : '') + ' name="chk' + i + '" style="display: none" type="checkbox" value="Y" /></td>' ;
		str += '<td align="right" width="60">' ;
		str += '<table border="0" cellPadding="0" cellSpacing="0">' ;
		str += '<tr>' ;
		str += '<td align="right"><input type="hidden" name="compRecQty' + i + '" value="' + RecQty + '">' ;
		str += '<input id="compQty' + i + '" class="input' + (LockQty || ItemCmd != 'A' ? 'Des' : '') + '" name="compQty' + i + '" onchange="compChangeField(1, ' + i + ');" onkeydown="return valKeyNumDec(event);" onfocus="this.select()" ' + (LockQty ? 'readOnly' : '') + ' size="12" style="TEXT-ALIGN: right; font-size: 8pt; font-face: Verdana;" value="' + Quantity + '" /><input type="hidden" id="compTreeType' + i + '" value="' + compTreeType + '">' ;
		str += '</td>' ;
		if (chkVolDisc) str += '<td style="height: 23px"><img src="images/foco_icon.gif" width="23" height="22" style="vertical-align: middle" onmouseover="javascript:showCompVolRep(this, event, ' + i + ');" onmouseout="javascript:clearVolRep();"></td>';
		str += '<td id="compInvErr' + i + '" ' + (chkInv || ItemCmd != 'A' ? 'style="display: none;"' : '') + '><img src="images/icon_alert.gif" alt="' + txtErrItmInv + '"></td>';
		str += '</tr>' ;
		str += '</table>' ;
		str += '</td>' ;
		
		if (itemDetailsCompType == 'ItemRec')
		{
			str += '<td style="width: 120px;">';
			
			if (!itemLockWhs && ItemCmd == 'A')
			{
				str += '<select id="compWhs' + i + '" size="1" class="input" style="font-face: Verdana; font-size: 8pt;" onchange="compChangeField(5, ' + i + ');">';
	
				for (var j = 0;j<arrWhs.length;j++)
				{
					var arrWhsData = arrWhs[j].split('{C}');
					str += '<option ' + (arrWhsData[0] == WhsCode ? 'selected' : '') + ' value="' + arrWhsData[0] + '">' + arrWhsData[1] + '</option>';
				}
				
				str += '</select>';
			}
			else
			{
				var strWhsName = '';
				
				for (var j = 0;j<arrWhs.length;j++)
				{
					var arrWhsData = arrWhs[j].split('{C}');
					if (arrWhsData[0] == WhsCode)
					{
						strWhsName = arrWhsData[1];
						break;
					}
				}
				
				str += '<input id="compWhs' + i + '" type="hidden" value="' + WhsCode + '"><span id="compWhsDesc' + i + '" style="font-weight: normal; font-size: xx-small;">' + strWhsName + '</span>';
			}
			
			str += '</td>';
		}
		else
		{
			str += '<input type="hidden" id="compWhs' + i + '" value="' + WhsCode + '">';
		}
		
		if (GetShowSalUn && itemDetailsCompType != 'SaleTree' && ItemCmd == 'A')
		{
			str += '<td align="middle"><select id="compUnit' + i + '" ' + (LockQty ? 'disabled' : '') + ' class="input' + (LockQty ? 'Des' : '') + '" onchange="javascript:compChangeField(2, ' + i + ');" style="width: 100px; font-face: Verdana; font-size: 8pt;">' +  
			'<option value="1">' + lblItemDetailsUnit + '</option>' +  
			'<option value="2" ' + (itemDetailsSaleType == 2 ? 'selected' : '') + '>' + SalUnitMsr + '(' + NumInSale + ')</option>' +  
			'<option value="3" ' + (itemDetailsSaleType == 3 ? 'selected' : '') + '>' + SalPackMsr + '(' + SalPackUn + ')</option>' +  
			'</select></td>';
		}
		else 
		{
			if (GetShowSalUn) str += '<td align="middle"><span class="CanastaTblResaltada" style="BACKGROUND-COLOR: #ffffff; FONT-WEIGHT: normal">' + SaleTypeDesc + '</span></td>' ;
			str += '<input type="hidden" id="compUnit' + i + '" value="' + compSaleType + '">';
		}
		
		str += '<input type="hidden" id="compPrevUnit' + i + '" value="' + itemDetailsSaleType + '">';
		
		if (enableItemLineDisc && ItemCmd == 'A')
		{
			str += '<td align="right">' ;
			str += '<input ' + (itemDetailsCompType == 'SaleTree' && compHideComp || olkCombo && !olkComboShowCompPrice ? 'type="hidden"' : '') + ' id="compDiscount' + i + '" class="input' + (LockPrice ? 'Des' : '') + '" ' + (LockPrice ? 'readOnly' : '') + ' name="compDiscount' + i + '" onchange="compChangeField(3, ' + i + ')" onkeydown="return valKeyNumDec(event);" onfocus="this.select()" size="8" style="TEXT-ALIGN: right; font-size: 8pt; font-face: Verdana;" value="%&nbsp;' + Discount + '">' ;
			str += '</td>';
		}
		else
		{
			str += '<input type="hidden" id="compDiscount' + i + '" name="compDiscount' + i + '" value="%&nbsp;' + Discount + '">' ;
		}
		str += '<td align="right">' ;
		str += '<input ' + (itemDetailsCompType == 'SaleTree' && compHideComp || olkCombo && !olkComboShowCompPrice || ItemCmd != 'A' ? 'type="hidden"' : '') + ' id="compPrice' + i + '" class="input' + (LockPrice ? 'Des' : '') + '" ' + (LockPrice ? 'readOnly' : '') + ' name="compPrice' + i + '" onchange="compChangeField(4, ' + i + ')" onkeydown="return valKeyNumDec(event);" onfocus="this.select()" size="18" style="TEXT-ALIGN: right; font-size: 8pt; font-face: Verdana;" value="' + Currency + '&nbsp;' + Price + '"><input type="hidden" id="compManPrc' + i + '" name="compManPrc' + i + '" value="N"></td>';
		str += '<td align="right"><input type="hidden" id="compHideComp' + i + '" value="' + (compHideComp ? 'Y' : 'N') + '">' ;
		str += '<input ' + (itemDetailsCompType == 'SaleTree' && compHideComp || olkCombo && !olkComboShowCompPrice || ItemCmd != 'A' ? 'type="hidden"' : '') + ' id="compLineTotal' + i + '" class="inputDes" readOnly name="compLineTotal' + i + '" onchange="" onkeydown="return valKeyNumDec(event);" onfocus="this.select()" size="18" style="TEXT-ALIGN: right; font-size: 8pt; font-face: Verdana;" value="' + DocCur + '&nbsp;' + LineTotal + '" /><input type="hidden" id="compCur' + i + '" value="' + Currency + '">' ;
		str += '</tr>' ;
		
	}
	
	$('#tbItemComp').empty();
	$('#tbItemComp').append(str);
	
	setItemDetailsBtnAddDis();

}

function GenerateItemComponents(itemRecData)
{
	var strItemComp = '<table border="0" cellPadding="0" width="100%">';
	strItemComp  += '<tr class="CanastaTblResaltada">';
	strItemComp  += '<td align="middle" style="WIDTH: 20px">&nbsp;</td>';
	strItemComp  += '<td align="middle">&nbsp;</td>';
	strItemComp  += '<td align="middle" colspan="2">' + lblItemDetailsQty + '</td>';
	
	strItemComp	 += '<td id="itemRecWhs" align="middle" style="width: 120px;">' + lblItemDetailsWhs + '</td>';
	if (GetShowSalUn) strItemComp += '<td align="middle" style="width: 120px;">' + lblItemDetailsUnit + '</td>';
	if (enableItemLineDisc)
	{
		strItemComp	 += '<td align="middle" style="width: 120px;">' + lblItemDetailsDisc + '</td>';
	}
	strItemComp  += '<td style="text-align: center; width: 108px;"><span id="spLblItmDetPrc">' + lblItemDetailsPrice + '</span></td>';
	strItemComp  += '<td style="text-align: center; width: 108px;"><span id="spLblItmDetTot">' + lblItemDetailsTotal  + '</span></td>';
	strItemComp  += '</tr>';
	strItemComp  += '<tBody id="tbItemComp"><tr><td>test</td></tr>';
	strItemComp  += '</tBody>';
	strItemComp  += '</table>';

	document.getElementById('itemDetTabs-2').innerHTML = strItemComp ;
	
	LoadItemComponents(itemRecData);
}

function LoadItemRep()
{
	if (enableItemRep)
	{
		var Quantity;
		var SaleType;
		var Price;
		if (ItemCmd == 'A') 
		{
			Quantity = document.getElementById('txtItemDetailsQty').value;
			SaleType = document.getElementById('txtItemDetailsSaleType').value;
			Price = document.getElementById('txtItemDetailsPrice').value.replace(itemDetailsCur, '');
		}
		$.post("Fetch/itemDetailsFetch.asp?d=" + (new Date()).toString(), { DataType: 'IR', Item: itemDetailsID, ItemCmd: ItemCmd, Quantity: Quantity, SaleType: SaleType, Price: Price },
		function(data)
		{
	   		var arrData = data.split('{S}');
	   		for (var i = 0;i<arrData.length - 1;i++)
	   		{
	   			var arrRepData = arrData[i].split('{C}');
	   			var txtItemRep = document.getElementById('itemRep' + arrRepData[0].replace('ItemRep', ''));
	   			
   				var doHide = document.getElementById('hideItemRep' + arrRepData[0].replace('ItemRep', ''));
	   			if (txtItemRep)
	   			{
	   				txtItemRep.innerHTML = arrRepData[1];
	   			}
   				if (doHide)
   				{
   					document.getElementById('itemRep' + arrRepData[0].replace('ItemRep', '') + 'Row').style.display = arrRepData[1] != '' ? '' : 'none';
   				}
	   		}
	   		
			rowComp = document.getElementById('itemRep_1Row');
			if (rowComp)
			{
				rowComp.style.display = itemDetailsTreeType == 'S' ? '' : 'none';
			}
			
			showItemDetailAJAXLoader('IR', false);
	   });
	}
}

function LoadItemInvRep()
{
	var WhsCode = document.getElementById('itemDetailsWhs').value;
	
	$.post("Fetch/itemDetailsFetch.asp?d=" + (new Date()).toString(), { DataType: 'Inv', Item: itemDetailsID, WhsCode: WhsCode },
	function(data)
	{
		var arrData = data.split('|');
		
		itemRep1LinkORDR = arrData[0] == 'Y';
		itemRep1LinkOBS = arrData[1] == 'Y';

		if (document.getElementById('txtItemDetailsInvOnHand'))
		{
			document.getElementById('txtItemDetailsInvOnHand').innerText = arrData[2];
			document.getElementById('txtItemDetailsInvOnHandUnVentSAP').innerText = arrData[3];
			document.getElementById('txtItemDetailsInvOnHandUnEmbSAP').innerText = arrData[4];
			
			document.getElementById('txtItemDetailsInvDispSAP').innerText = arrData[5];
			document.getElementById('txtItemDetailsInvDispUnVentSAP').innerText = arrData[6];
			document.getElementById('txtItemDetailsInvDispUnEmbSAP').innerText = arrData[7];
		}
		
		
		if (document.getElementById('txtItemDetailsInvInvBDGWhs'))
		{
			document.getElementById('txtItemDetailsInvInvBDGWhs').innerText = arrData[8];
			document.getElementById('txtItemDetailsInvInvUnVentBDGWhs').innerText = arrData[9];
			document.getElementById('txtItemDetailsInvInvUnEmbBDGWhs').innerText = arrData[10];
			
			document.getElementById('txtItemDetailsInvInvBDGDisp').innerText = arrData[11];
			document.getElementById('txtItemDetailsInvInvBDGUnVentDisp').innerText = arrData[12];
			document.getElementById('txtItemDetailsInvInvBDGUnEmbDisp').innerText = arrData[13];
		}
		
		if (document.getElementById('txtItemDetailsInvInvOLKDisp'))
		{
			document.getElementById('txtItemDetailsInvInvOLKDisp').innerText = arrData[14];
			document.getElementById('txtItemDetailsInvInvOLKUnVentDisp').innerText = arrData[15];
			document.getElementById('txtItemDetailsInvInvOLKUnEmbDisp').innerText = arrData[16];
			
			document.getElementById('txtItemDetailsInvInvOLKBDGDisp').innerText = arrData[17];
			document.getElementById('txtItemDetailsInvInvOLKBDGUnVentDisp').innerText = arrData[18];
			document.getElementById('txtItemDetailsInvInvOLKBDGUnEmbDisp').innerText = arrData[19];
		}
		
		if (document.getElementById('itemRep1LinkORDR')) document.getElementById('itemRep1LinkORDR').style.display = itemRep1LinkORDR ? '' : 'none';
		if (document.getElementById('itemRep1LinkOBS')) document.getElementById('itemRep1LinkOBS').style.display = itemRep1LinkOBS ? '' : 'none';
		
		if (document.getElementById('itemRep1LinkORDRSpan'))
			if (!itemRep1LinkORDR) document.getElementById('itemRep1LinkORDRSpan').colSpan = 2;
			else document.getElementById('itemRep1LinkORDRSpan').colSpan = 1;
		
		if (document.getElementById('itemRep1LinkOBSSpan'))
			if (!itemRep1LinkOBS) document.getElementById('itemRep1LinkOBSSpan').colSpan = 2;
			else document.getElementById('itemRep1LinkOBSSpan').colSpan = 1;
		
		showItemDetailAJAXLoader('Inv', false);
	});
}

function LoadItemDetails(dv, item)
{
	$.post("Fetch/itemDetailsFetch.asp?d=" + (new Date()).toString(), { DataType: 'L', ItemCmd: ItemCmd },
   function(data)
   {
   		var arrData = data.split('{S}');
		arrWhs = arrData[0].split('{O}');
		
		enableItemRep = arrData[1] != '';
		if (enableItemRep) arrItemRep = arrData[1].split('{O}');
		
		enableItemInvRep = arrData[2] != '';
		var arrInv;
		if (enableItemInvRep) arrInv = arrData[2].split('{C}');
		
		
		vDisp = arrData[3];
		
		enableItemLineDisc = arrData[4] == 'Y';
		
		var predefinedNotes = arrData[5] != '';
		var arrNotes;
		if (predefinedNotes)
		{
			arrNotes = arrData[5].split('{N}');
		}
		
		var SDKLineMemo = arrData[6] == 'Y';
		itemLockWhs = arrData[7] == 'Y';
		
		var taxCodeData = arrData[8];

		var str = '<div style="border: 1px solid #C1E3FF;"><table border="0" cellpadding="0" width="100%">' +  
			'<tr>' +  
			'<td>' +  
			'<table border="0" cellpadding="0" width="100%">' +  
			'<tr class="GeneralTbl">' +  
			'<td height="100" width="15%">' +  
			'<p align="center"><a href="#" onclick="javascript:itemDetailsOpenPic()">' +  
			'<img id="itemDetailsImg" src="pic.aspx?filename=n_a.gif&dbName=' + dbName + '" border="0"></a></p>' +  
			'</td>' +  
			'<td colspan="2" valign="top">' +  
			'<ilayer name="scroll1" width=100% height=100 clip="0,0,170,150">' +  
			'<layer bgcolor="white" height="100" name="scroll2" width="100%">' +  
			'<div id="scroll3" style="width: 100%; height: 100px; background-color: white; overflow: auto">' +  
			'<span id="itemDetailsText"></span></div>' +  
			'</layer>' +  
			'</ilayer>' +  
			'</td>' +  
			'</tr>' +  
			'</table>' +  
			'</td>' +  
			'</tr>' +  
			'</table>' + 
			'<div id="tblItemDetailsTop" class="Transparency" style="display: none;position: absolute; left: 0px; top: 0px; filter:alpha(opacity=60); background: rgb(225,238,253); height: 148px; width: 100%;"></div>' +
			'<img src="design/0/images/ajax_popup_loader.gif" id="tblItemDetailsTopImg" style="display: none;position: absolute; top: 40px; left: 370px;">' +
			'</div>' +
			'<div id="itemDetTabs" style="border: 1px solid #C1E3FF;">' + 
			'<ul>' +
			'<li><a href="#itemDetTabs-1" style="font-size: xx-small;">' + txtDetails + '</a></li>' +
			'<li id="liComponents"><a href="#itemDetTabs-2" style="font-size: xx-small;">' + txtComponents + '</a></li>';
			
			if (itemRepLinkInvRep) str += '<li><a href="#itemDetTabs-3" style="font-size: xx-small;">' + txtInv + '</a></li>';
			
			if (itemRepLinkBestPrice) str += '<li><a href="#itemDetTabs-4" style="font-size: xx-small;">' + txtBestPrices + '</a></li>';
			
			if (itemRepLinkLastSale) str += '<li><a href="#itemDetTabs-5" style="font-size: xx-small;">' + txtSaleRep + '</a></li>';
			
			if (itemRep1LinkORDR) str += '<li><a href="#itemDetTabs-6" style="font-size: xx-small;">' + txtOLKCommited + '</a></li>';
			
			if (itemRep1LinkOBS) str += '<li><a href="#itemDetTabs-7" style="font-size: xx-small;">' + txtSAPCommited + '</a></li>';
			
			str += '</ul>' +
			'<div id="itemDetTabs-1" style="height: 258px; background-color: #FFFFFF; overflow: auto;"><table border="0" cellpadding="0" width="100%">' +   
			'<tr>' +  
			'<td>' +  
			'<table border="0" cellpadding="0" cellspacing="1" width="100%">' +  
			'<tr>' +  
			'<td rowspan="2" valign="top" width="420" style="border: 1px solid #C1E3FF;">' +  
			'<div id="itemDetailsLeft" style="margin: 1px; height: 240px; width: 100%; overflow: auto; overflow-x: none; overflow-y: scroll; background-color: #FFFFFF;"> ' + 
							'<table border="0" cellpadding="0" cellspacing="1" width="100%">';
				
			if (enableItemRep)
			{			
				for (var i = 0;i<arrItemRep.length;i++)
				{
					var arrItemRepData = arrItemRep[i].split('{C}');
					var repIndex = arrItemRepData[0];
					
					str += '<tr id="itemRep' + arrItemRepData[0].replace('-', '_') + 'Row"><td style="vertical-align: top;" class="GeneralTblBold2' + (repIndex == '-1' ? 'HighLight' : '') + '">';
					str += '<table cellpadding="0" cellspacing="0" border="0" width="100%; height: 100%; "><tr class="GeneralTblBold2' + (repIndex == '-1' ? 'HighLight' : '') + '">';
					str += '<td>' + arrItemRepData[1] + '</td>'
					if (arrItemRepData[2] == 'Y') 
					{
						str += '<td align="' + (rtl == '' ? 'right' : 'left') + '">';
						str += '<img alt="' + arrItemRepData[3] + '" border="0" src="design/0/images/' + rtl + 'felcahSelect.gif" width="15" height="13" style="cursor: hand" onclick="javascript:doItemRepLink(' + arrItemRepData[0] + ', ' + arrItemRepData[3] + ');"></td>';
					}
					str += '</tr></table></td><td class="Generaltbl" id="itemRep' + arrItemRepData[0].replace('-', '_') + '">';

					if (arrItemRepData[5] == 'Y')  str += '<input type="hidden" id="hideItemRep' + arrItemRepData[0].replace('-', '_') + '" value="Y">';
					
					str += '&nbsp;</td></tr>';
				}
			}

			str += '</table>' +
			'<div id="tblItemDetailsIR" class="Transparency" style="position: absolute; left: 0px; top: 0px; filter:alpha(opacity=60); background: rgb(225,238,253); height: 232px; width: 100%;"></div>' +
			'<img src="design/0/images/ajax_popup_loader.gif" id="tblItemDetailsIRImg" style="position: absolute; top: 80px; left: 105px;">' +
			'</div>' +  
			'</td>' +  
			'<td valign="top" style="border: 1px solid #C1E3FF;">' +  
			'<p align="center"></p>' +  
			'</td>' +  
			'</tr>' +  
			'<tr>' +  
			'<td valign="top" style="border: 1px solid #C1E3FF;">' +  
			'<div id="itemDetailsInv" style="margin: 1px; height: 200px; width: 100%; border: 0px; overflow: auto; overflow-x: none;">';
			

			if (enableItemInvRep) 
			{
				str += '<table border="0" cellpadding="0" width="100%">' +  
				'<tr class="GeneralTblBold2">' +  
				'<td colspan="2">&nbsp;</td>' +  
				'<td>' + lblItemDetailsUnit + '</td>' +  
				'<td>' + txtSalUn + '</td>' +  
				'<td>' + txtPackUn + '</td>' +  
				'</tr>' +  
				'<tr class="GeneralTblBold2">' +  
				'<td colspan="2">&nbsp;</td>' +  
				'<td>' + lblItemDetailsUnit + '</td>' +  
				'<td id="txtItemDetailsInvUnit2"></td>' +  
				'<td id="txtItemDetailsInvUnit3"></td>' +  
				'</tr>';
				for (var i = 0;i<arrInv.length;i++)
				{
					switch (arrInv[i])
					{
						case 'SAP':
							str +=  '<tr>' + 
							'<td colspan="2" class="GeneralTblBold2">' + txtSAP + '</td>' + 
								'<td class="Generaltbl" align="right" dir="ltr" id="txtItemDetailsInvOnHand"></td>' + 
								'<td class="Generaltbl" align="right" dir="ltr" id="txtItemDetailsInvOnHandUnVentSAP"></td>' + 
								'<td class="Generaltbl" align="right" dir="ltr" id="txtItemDetailsInvOnHandUnEmbSAP"></td>' + 
								'</tr>' + 
								'<tr>' + 
								'<td id="itemRep1LinkORDRSpan" class="GeneralTblBold2' + (vDisp == 'SD' ? 'HighLight' : '') + '">' + txtAVL + '</td>' + 
								'<td id="itemRep1LinkORDR" style="display: none; width: 15px;" class="GeneralTblBold2' + (vDisp == 'SD' ? 'HighLight' : '') + '">' + 
								'<a href="#" onclick="javascript:goItemDetTab(6);"><img border="0" src="design/0/images/' + rtl + 'felcahSelect.gif" width="15" height="13"></a></td>' +
								'<td class="Generaltbl" align="right" dir="ltr" id="txtItemDetailsInvDispSAP"></td>' + 
								'<td class="Generaltbl" align="right" dir="ltr" id="txtItemDetailsInvDispUnVentSAP"></td>' + 
								'<td class="Generaltbl" align="right" dir="ltr" id="txtItemDetailsInvDispUnEmbSAP"></td>' + 
								'</tr>';
							break;
						case 'OLK':
							str += 	'<tr>' +  
									'<td id="itemRep1LinkOBSSpan" class="GeneralTblBold2' + (vDisp == 'OD' ? 'HighLight' : '') + '">' + txtOLK + '</td>' + 
									'<td id="itemRep1LinkOBS" style="display: none; width: 15px;" class="GeneralTblBold' + (vDisp == 'OD' ? 'HighLight' : '') + '">' +  
									'<a href="#" onclick="javascript:goItemDetTab(5);"><img border="0" src="design/0/images/' + rtl + 'felcahSelect.gif" width="15" height="13"></a></td>' + 
									'<td class="Generaltbl" align="right" dir="ltr" id="txtItemDetailsInvInvOLKDisp">&nbsp;</td>' +  
									'<td class="Generaltbl" align="right" dir="ltr" id="txtItemDetailsInvInvOLKUnVentDisp"></td>' +  
									'<td class="Generaltbl" align="right" dir="ltr" id="txtItemDetailsInvInvOLKUnEmbDisp"></td>' +  
									'</tr>' +  
									'<tr>' +  
									'<td class="GeneralTblBold2' + (vDisp == 'OS' || vDisp == 'OE' ? 'HighLight' : '') + '" colspan="2">' + txtAVL + '</td>' +  
									'<td class="Generaltbl" align="right" dir="ltr" id="txtItemDetailsInvInvOLKBDGDisp">&nbsp;</td>' +  
									'<td class="Generaltbl" align="right" dir="ltr" id="txtItemDetailsInvInvOLKBDGUnVentDisp"></td>' +  
									'<td class="Generaltbl" align="right" dir="ltr" id="txtItemDetailsInvInvOLKBDGUnEmbDisp"></td>' +  
									'</tr>' 
							break;
						case 'BDG':
							str += '<tr>' +  
								'<td class="GeneralTblBold2" colspan="2">' + txtWHS + '</td>' +  
								'<td class="Generaltbl" align="right" dir="ltr" id="txtItemDetailsInvInvBDGWhs">&nbsp;</td>' +  
								'<td class="Generaltbl" align="right" dir="ltr" id="txtItemDetailsInvInvUnVentBDGWhs"></td>' +  
								'<td class="Generaltbl" align="right" dir="ltr" id="txtItemDetailsInvInvUnEmbBDGWhs"></td>' +  
								'</tr>' +  
								'<tr>' +  
								'<td class="GeneralTblBold2' + (vDisp == 'SE' || vDisp == 'SS' ? 'HighLight' : '') + '" colspan="2">' + txtAVL + '</td>' +  
								'<td class="Generaltbl" align="right" dir="ltr" id="txtItemDetailsInvInvBDGDisp">&nbsp;</td>' +  
								'<td class="Generaltbl" align="right" dir="ltr" id="txtItemDetailsInvInvBDGUnVentDisp"></td>' +  
								'<td class="Generaltbl" align="right" dir="ltr" id="txtItemDetailsInvInvBDGUnEmbDisp"></td>' +  
								'</tr>';
							break;
					}
				}
				str += '</table>';
			}
			 
			str += '<div id="tblItemDetailsInv" class="Transparency" style="position: absolute; left: 0px; top: 0px; filter:alpha(opacity=60); background: rgb(225,238,253); height: 200px; width: 100%;"></div>' +
			'<img src="design/0/images/ajax_popup_loader.gif" id="tblItemDetailsInvImg" style="position: absolute; top: 90px; left: 210px;">' +
			'</div>' +  
			'</td>' +  
			'</tr>' +  
			'</table>' +  
			'</td>' +  
			'</tr>' + 
			'</table></div>' +
			'<div id="itemDetTabs-2" style="height: 258px; background-color: #FFFFFF; overflow: auto;"></div>';
			
			if (itemRepLinkInvRep) str += '<div id="itemDetTabs-3" style="height: 258px; background-color: #FFFFFF; overflow: auto;"></div>';
			
			if (itemRepLinkBestPrice)
			{
				str += '<div id="itemDetTabs-4" style="height: 258px; background-color: #FFFFFF; overflow: auto;">' +
				'<table border="0" cellpadding="0" width="100%">' +  
				'<tr>' +  
				'<td colspan="5">' +  
				'<table cellpadding="0" cellspacing="0" width="100%">' +  
				'<tr>' +  
				'<td class="GeneralTblBold2">' + lblItemDetailsBestPrice + ':</td>' +  
				'<td class="GeneralTbl"><nobr><span id="txtItemDetailsBestPrice"></span></nobr></td>' +  
				'<td class="GeneralTblBold2">' + lblItemDetailsQty + ':</td>' +  
				'<td class="GeneralTbl"><span id="txtItemDetailsBestQty"></span></td>' +  
				'</tr>' +  
				'</table>' +  
				'</td>' +  
				'</tr>' +  
				'<tr class="GeneralTblBold2">' +  
				'<td align="center" colspan="2">' + txtInvoice + '</td>' +  
				'<td align="center">' + lblItemDetailsDate + '</td>' +  
				'<td align="center">' + lblItemDetailsQty + '</td>' +  
				'<td align="center">' + lblItemDetailsPrice + '</td>' +  
				'</tr>' +  
				'<tBody id="tblItemDetailsBestPrice">' +  
				'</tBody>' +  
				'<tr class="GeneralTbl">' +  
				'<td>&nbsp;</td>' +  
				'<td>&nbsp;</td>' +  
				'<td>&nbsp;</td>' +  
				'<td>&nbsp;</td>' +  
				'<td>&nbsp;</td>' +  
				'</tr>' +  
				'</table>' +
				'</div>';
			}
			
			if (itemRepLinkLastSale) str += '<div id="itemDetTabs-5" style="height: 258px; background-color: #FFFFFF; overflow: auto;"></div>';
			
			if (itemRep1LinkORDR) str += '<div id="itemDetTabs-6" style="height: 258px; background-color: #FFFFFF; overflow: auto;"></div>';
			
			if (itemRep1LinkOBS) str += '<div id="itemDetTabs-7" style="height: 258px; background-color: #FFFFFF; overflow: auto;"></div>';
			
			str += '</div>' +
			'<div id="tblItemDetailsWhole" class="Transparency" style="display: none; position: absolute; left: 0px; top: 140px; filter:alpha(opacity=60); background: rgb(225,238,253); height: 264px; width: 100%;"></div>' +
			'<img src="design/0/images/ajax_popup_loader.gif" id="tblItemDetailsWholeImg" style="display: none; position: absolute; top: 250px; left: 440px;">' +
			'</div>' +  
			'<div style="border: 1px solid #C1E3FF;"><table border="0" cellpadding="0" width="100%">' +   
			'<tr>' +  
			'<td>' +  
			'<table border="0" cellpadding="0" cellspacing="0" width="100%">' +  
			'<tr>' +  
			'<td>' +  
			'<table border="0" cellpadding="0" cellspacing="0" width="100%">' +  
			'<tr id="itemDetailsMsgRow" class="GeneralTblBold2" style="height: 22px; text-align: center;">' +  
			'<td>';
			
			str += '<table cellpadding="0" cellspacing="0" border="0" id="tblVolDisc" style="display: none;">' +  
					'<tr class="GeneralTblBold2" style="background-color:#CCFF99; text-align: center; ">' +  
					'<td>' +  
					'<img src="images/foco_icon.gif" width="23" height="22" style="vertical-align: middle" onmouseover="javascript:showItemVolRep(this, event);" onmouseout="javascript:clearVolRep();">' +  
					'</td><td>' + txtVolDiscAvl + '</td></tr></table>' + 
					'<table cellpadding="0" border="0" id="tblItemDetailsMsg">' +  
					'<tr>';
					
			if (EnableCartSum)
			{
				str += '<td width="15" id="tdItemDetailsCartFilterLink" style="cursor: pointer;" onclick="javascript:itemDetailsFilterCart();"><img border="0" src="design/0/images/' + rtl + 'felcahSelect.gif" width="15" height="13" alt="' + txtFilterInCart + '"></td>';
			}
			
			str += '<td class="GeneralTblBold2" style="background-color:#FFD2A6;" id="itemDetailsMsg"></td>' +  
					'</tr>' +  
					'</table><span id="lblSalesHideCompTxt" style="display: none;">' + txtSalesHideCompTxt + '</span>';
			
			str += '</td>' +  
			'</tr>' +  
			'</table>';
			
			switch (ItemCmd)
			{
				case 'A':
				case 'W':
			
					str += '<table cellpadding="0" width="100%" id="tblItemDetailsControlsAI1">' +  
					'<tr>' +  
					'<td class="GeneralTblBold2" style="text-align: center;">' +  
					'' + lblItemDetailsQty + '</td>' + 
					'<td class="GeneralTblBold2" style="text-align: center;" id="ttlCartQty">' +  
					'' + txtCartQty + '</td>' +  
					'<td class="GeneralTblBold2" style="text-align: center;">' +  
					'' + lblItemDetailsWhs  + '</td>' + 
					'<td class="GeneralTblBold2" style="text-align: center;">' +  
					'' + lblItemDetailsUnit + '</td>';
					
					if (enableItemLineDisc)
					{
						str += '<td class="GeneralTblBold2" style="text-align: center;">' + lblItemDetailsDisc + '</td>';
					}
					
					str += '<td class="GeneralTblBold2" style="width: 120px; text-align: center;">' +  
					'<span id="spItemDetailsPrice">' + lblItemDetailsPrice + '</span></td>';
					 
					if (LawsSet == 'MX' || LawsSet == 'CL' || LawsSet == 'CR' || LawsSet == 'GT' || LawsSet == 'US' || LawsSet == 'CA')
					{
						str += '<td style="text-align: center;" class="GeneralTblBold2">' + txtTaxCode + '</td>';
					}
					
					
					str += '<td class="GeneralTblBold2" colspan="2" style="width: 120px; text-align: center;">' + 
					'<span id="spItemDetailsTotal">' + lblItemDetailsTotal + '</span></td>' +  
					'</tr>' +  
					'<tr>' +  
					'<td class="Generaltbl" style="text-align: center;">' +  
					'<table cellpadding="0" cellspacing="0" border="0"><tr><td class="Generaltbl"><input id="txtItemDetailsQty" onchange="itemDetailsChangeField(1);" onfocus="this.select()" onkeydown="return valKeyNumDec(event);" size="22" style="text-align: right" type="text" value="1.00"></td><td id="ItemDetailsInvErr" style="display: none;"><img src="images/icon_alert.gif" alt="' + txtErrItmInv + '"></td></tr></table></td>' +  
					'<td class="Generaltbl" id="tdCartQty" style="text-align: center;">' +  
					'<input id="txtItemDetailsCartQty" disabled class="InputDes" onfocus="this.select()" size="22" style="text-align: right" type="text" value="1.00"></td>' +  
					'<td class="Generaltbl" style="text-align: center;">';
					
		
					if (!itemLockWhs)
					{
						str += '<select id="itemDetailsWhs" size="1" onchange="javascript:itemDetailsChangeField(5);">';
						
						for (var i = 0;i<arrWhs.length;i++)
						{
							var arrWhsData = arrWhs[i].split('{C}');
							str += '<option value="' + arrWhsData[0] + '">' + arrWhsData[1] + '</option>';
						}
						
						str += '</select>';
					}
					else
					{
						str += '<input id="itemDetailsWhs" type="hidden" value=""><span id="itemWhsDesc" style="font-weight: normal; font-size: xx-small;"></span>';
					}
					
					str += '</td>' +
					'<td class="Generaltbl" style="text-align: center;">' +  
					'<select id="txtItemDetailsSaleType" onchange="javascript:itemDetailsChangeField(2);" style="width: 100px">' +  
					'<option value="1">' + lblItemDetailsUnit + '</option>' +  
					'<option id="optItemDetails2" value="2">UND(1)' +  
					'</option>' +  
					'<option id="optItemDetails3" value="3">UND(1)</option>' +  
					'</select></td>';
					
					if (enableItemLineDisc)
					{
						str += '<td class="Generaltbl" style="text-align: center;">' +  
						'<input id="txtItemDetailsDisc" class="input" onchange="javascript:itemDetailsChangeField(3);" onkeydown="return valKeyNumDec(event);" onfocus="this.select()" style="width: 60px; text-align: right" type="text" value="%&nbsp;0.000">' +  
						'</td>';
					}
					else
					{
						str += '<input id="txtItemDetailsDisc" type="hidden" value="%&nbsp;0.000">';
					}
					
					str += '<td class="Generaltbl">' +  
					'<input id="txtItemDetailsPrice" class="input" onchange="javascript:itemDetailsChangeField(4);" onkeydown="return valKeyNumDec(event);" onfocus="this.select()" style="width: 120px; text-align: right" type="text" value="0.00">' +  
					'</td>';
					
					if (LawsSet == 'MX' || LawsSet == 'CL' || LawsSet == 'CR' || LawsSet == 'GT' || LawsSet == 'US' || LawsSet == 'CA')
					{
						str += '<td class="Generaltbl"><select class="input" name="TaxCode" id="TaxCode" style="width:98%">' +
								'<option></option>';
								
						if (taxCodeData == '')
						{
							str += '<option value="">' + txtNotApply + '</option>';
						}
						else
						{
							var arrTaxCode = taxCodeData.split('{C}');
							for (var d = 0;d<arrTaxCode.length;d++)
							{
								var arrTaxCodeData = arrTaxCode[d].split('{D}');
								str += '<option value="' + arrTaxCodeData[0] + '">' + arrTaxCodeData[0] + ' - ' + arrTaxCodeData[1] + '</option>';
							}
						}
						
						str += '</select></td>';
					}
					
					str += '<td class="Generaltbl" colspan="2" style="text-align: right;">' +  
					'<input id="txtItemDetailsTotal" class="InputDes" onfocus="this.select()" readonly style="width: 120px; text-align: right" type="text"></td>' +  
					'</tr>' +  
					'</table>' +  
					'<table cellpadding="0" class="style1" width="100%" id="tblItemDetailsControlsAI2">' +  
					'<tr>' +  
					'<td class="GeneralTblBold2" style="text-align: center; width: 100px;">' +  
					'' + lblItemDetailsNote + '</td>' +  
					'<td class="GeneralTblBold2" rowspan="2" valign="top">' +  
					'<textarea id="txtItemDetailsMemo" ' + (!SDKLineMemo ? 'disabled' : '') + ' onkeydown="return chkMax(event, this, 254);" style="width: 100%">' + (!SDKLineMemo ? txtDisNotes : '') + '</textarea></td>' +  
					'</tr>' +  
					'<tr>' +  
					'<td class="GeneralTblBold2" width="100">' +  
					'<select id="txtItemDetailsCmbNote" class="input" ' + (!SDKLineMemo ? 'disabled' : '') + ' onchange="document.getElementById(\'txtItemDetailsMemo\').value=this.value;" size="1" style="width: 100px">' +  
					'<option value="">' + lblItemDetailsChooseFrom + '</option>';
					
					if (predefinedNotes && SDKLineMemo)
					{
						for (var i = 0;i<arrNotes.length;i++)
						{
							var arrNote = arrNotes[i].split('{C}');
							str += '<option value="' + arrNote[1] + '">' + arrNote[0] + '</option>';
						}
					}
					
					str += '</select></td>' +  
					'</tr>' +  
					'</table>' +  
					'<table cellpadding="0" cellspacing="0" style="border-width: 0px; border-style: solid;" width="100%" id="tblItemDetailsControlsAI3">' +  
					'<tr class="GeneralTblBold2">' +  
					'<td>' +  
					'<table border="0" cellpadding="0" cellspacing="0" width="100%">' +  
					'<tr class="GeneralTblBold2">' +  
					'<td style="width: 50px">' +  
					'<input id="txtItemDetailsBtnAdd" type="button" value="' + lblItemDetailsConfirm + '" onclick="javascript:itemDetailsAdd();"></td>';

					if (EnSellAll && ItemCmd == 'A')
					{
						str += '<td>&nbsp;<input type="checkbox" name="txtItemDetailsAddAll" value="Y" id="txtItemDetailsAddAll" style="border-style:solid; border-width:0px;" onclick="javascript:enableItemDetailsQty(this.checked);"><label for="txtItemDetailsAddAll">' + txtSellAll + '</label></td>';
					}
					
					str += '<td style="text-align: right">&nbsp;<input id="txtItemDetailsBtnCancel" type="button" value="' + lblItemDetailsCancel + '" onclick="javascript:closeItemDetails();"></td>' +  
					'</tr>' +  
					'</table>' +  
					'</td>' +  
					'</tr>' +  
					'</table>';
					
					str += '<table cellpadding="0" cellspacing="0" style="border-width: 0px; border-style: solid; display: none;" width="100%" height="100" id="tblItemDetailsControlsWL">' +  
					'<tr class="GeneralTblBold2">' +  
					'<td valign="bottom">' +  
					'<table border="0" cellpadding="0" cellspacing="0" width="100%">' +  
					'<tr class="GeneralTblBold2">' +
					'<td style="width: 50px">' +
					'<input id="txtItemDetailsBtnAdd" type="button" value="' + lblItemDetailsConfirm + '" onclick="javascript:itemDetailsAddWL();"></td>' + 
					'<td style="text-align: ' + (rtl == '' ? 'right' : 'left') + '"><input id="txtItemDetailsBtnCancel" type="button" value="' + (ItemCmd == 'W' ? lblItemDetailsCancel : lblItemDetailsClose) + '" onclick="javascript:closeItemDetails();"></td>' +  
					'</tr>' +  
					'</table>' +  
					'</td>' +  
					'</tr>' +  
					'</table>';

					break;
				case 'D':
					str += '<input id="itemDetailsWhs" type="hidden" value="">' + 
					'<table cellpadding="0" cellspacing="0" style="border-width: 0px; border-style: solid;" width="100%" height="100">' +  
					'<tr class="GeneralTblBold2">' +  
					'<td valign="bottom">' +  
					'<table border="0" cellpadding="0" cellspacing="0" width="100%">' +  
					'<tr class="GeneralTblBold2">';
					
					if (ItemCmd == 'W')
					{
						str += '<td style="width: 50px">' +
						'<input id="txtItemDetailsBtnAdd" type="button" value="' + lblItemDetailsConfirm + '" onclick="javascript:itemDetailsAddWL();"></td>';
					}
					
					str += '<td style="text-align: ' + (rtl == '' ? 'right' : 'left') + '"><input id="txtItemDetailsBtnCancel" type="button" value="' + (ItemCmd == 'W' ? lblItemDetailsCancel : lblItemDetailsClose) + '" onclick="javascript:closeItemDetails();"></td>' +  
					'</tr>' +  
					'</table>' +  
					'</td>' +  
					'</tr>' +  
					'</table>';

					break;
			}  
			
			str += '</td>' +  
			'</tr>' +  
			'</table>' +  
			'</td>' +  
			'</tr>' +   
			'</table>' +
			'<div id="tblItemDetailsBottom" class="Transparency" style="display: none; position: absolute; left: 1px; top: 396px; filter:alpha(opacity=60); background: rgb(225,238,253); height: 122px; width: 100%;"></div>' +
			'<img src="design/0/images/ajax_popup_loader.gif" id="tblItemDetailsBottomImg" style="display: none; position: absolute; top: 420px; left: 370px;">' +
			'</div>';


		dv.innerHTML = str;
		
		$('#itemDetTabs').tabs( { select: function(event, ui) 
											{ 
												switch (ui.panel.id)
												{
													case 'itemDetTabs-3':
														LoadItemDetailsBdgDetailsRep();
														break;
													case 'itemDetTabs-4':
														LoadItemDetailsBP();
														break;
													case 'itemDetTabs-5':
														LoadItemDetailsSalesRep();
														break;
													case 'itemDetTabs-6':
														LoadItemDetailsCommRep('O');
														break;
													case 'itemDetTabs-7':
														LoadItemDetailsCommRep('S');
														break;
												}
											 } } );
		
		LoadItemDetailsData(item);
	});
}

function enableItemDetailsQty(chk)
{
	txtItemDetailsQty.disabled = chk;
}
function doItemChangeField(id, isComp, itemCode, txtItemDetailsQty, TreeType, SaleType, PrevSaleType, txtItemDetailsDisc, txtItemDetailsPrice, txtItemDetailsTotal, imgInvErr, itemDetailsWhs, DiscPrcnt, ManPrc, Currency, compID)
{
	if (!isComp)
	{
		if (id == 5)
			if (enableItemInvRep)
			{
				showItemDetailAJAXLoader('Bottom', true);
			}	
		
		if (enableItemRep)
		{
			showItemDetailAJAXLoader('IR', true);
		}
	}

	var price = txtItemDetailsPrice.value;
	price = price.replace(Currency, '');
	
	if (!isComp && olkCombo && olkComboVirtual && id == 4)
	{
		DiscPrcnt = 100-parseFloat(getNumericVB(price))*100/parseFloat(getNumericVB(itemDetailsFatherUnitPrice));
	}

	$.post("Fetch/itemDetailsFetch.asp?d=" + (new Date()).toString(), { DataType: 'CF', 
																		Item: itemCode, 
																		ManPrc: ManPrc, 
																		SaleType: SaleType,
																		PrevSaleType: PrevSaleType, 
																		Price: price,
																		DiscPrcnt: DiscPrcnt,
																		Quantity: txtItemDetailsQty.value,
																		WhsCode: itemDetailsWhs.value,
																		FieldID: id,
																		TreeType: TreeType,
																		OLKCombo: (olkCombo ? 'Y' : 'N'),
																		isComp: (isComp ? 'Y' : 'N'),
																		ShowComp: (olkComboShowComp ? 'Y' : 'N'),
																		Virtual: (olkComboVirtual ? 'Y' : 'N'),
																		cmd: 'A',
																		Currency: Currency
																		},function(data)
	{
		var arrData = data.split('{S}');
		var qty = arrData[0];
		var price = arrData[1];
		var discPrcnt = arrData[2];
		var lineTotal = arrData[3];
		var chkInv = arrData[4] == 'Y';
		var cur = arrData[5];
		var itemRecData = arrData[6];
		var virtualTotal = arrData[7];
		var virtualChkInv = arrData[8];
		var virtualDiscPrcnt = arrData[9];
		var virtualPrice = arrData[10];
		
		if (olkCombo && olkComboVirtual && !isComp)
		{
			lineTotal = virtualTotal;
			chkInv = virtualChkInv;
			price = virtualPrice;
			discPrcnt = virtualDiscPrcnt;
		}
		
		txtItemDetailsQty.value = qty;
		txtItemDetailsPrice.value = cur + ' ' + price;
		if (txtItemDetailsDisc) txtItemDetailsDisc.value = '% ' + discPrcnt;
		txtItemDetailsTotal.value = cur + ' ' + lineTotal;
		
		
		imgInvErr.style.display = chkInv ? 'none' : '';
		
		if (id == 2) itemDetailsSaleType = txtItemDetailsSaleType.value;
		
		if ((id == 1 || id == 2 || olkCombo && olkComboVirtual) && itemDetailsCompType != '' && !isComp) 
		{
			LoadItemComponents(itemRecData);
		}
		showItemDetailAJAXLoader('Whole', false);

		if (!isComp)
		{
			if (id == 5)
				if (enableItemInvRep)
				{
					LoadItemInvRep();
				}	
			
			if (enableItemRep)
			{
				LoadItemRep();
			}
		}	
		else
		{
			var chkComp = document.getElementById('chkComp' + compID);
			if (!chkComp.checked)
			{
				doCheckComp(compID);
			}
			else if (olkCombo && olkComboVirtual)
			{
				SumChildTotal();
			}
		}
		setItemDetailsBtnAddDis();
		showItemDetailAJAXLoader('Bottom', false);
	});
}

function SumChildTotal()
{
	var compData = '';
	
	for (var i = 0;i<itemDetailsCompCount;i++)
	{
		var itemCode = document.getElementById('compItem' + i);
		var compQty = document.getElementById('compQty' + i);
		var compPrice = document.getElementById('compPrice' + i);
		var compUnit = document.getElementById('compUnit' + i);
		var compCur = document.getElementById('compCur' + i);
		var compLineTotal = document.getElementById('compLineTotal' + i);
		var chkComp = document.getElementById('chkComp' + i);

		if (chkComp.checked)
		{
			if (compData != '') compData += '{I}';
			compData += compPrice.value.replace(compCur.value, '') + '{S}' + compCur.value + '{S}' + compQty.value + '{S}' + compUnit.value + '{S}' + itemCode.value;
		}
	}

	$.post("Fetch/itemDetailsFetch.asp?d=" + (new Date()).toString(), { DataType: 'CT', 
																		Currency: itemDetailsCur, 
																		CompData: compData,
																		Quantity: document.getElementById('txtItemDetailsQty').value,
																		DiscPrcnt: document.getElementById('txtItemDetailsDisc').value.replace('%', '')
																		},function(data)
																		{
																			var arrData = data.split('{S}');
																			itemDetailsManPrc = true;
																			
																			document.getElementById('txtItemDetailsTotal').value = itemDetailsCur + ' ' + arrData[0];
																			document.getElementById('txtItemDetailsPrice').value = itemDetailsCur + ' ' + arrData[1];
																			
																			itemDetailsFatherUnitPrice = arrData[2];
																		});
}

function itemDetailsChangeField(id)
{
	showItemDetailAJAXLoader('Bottom', true);
	if ((id == 1 || id == 2) && itemDetailsCompType != '') showItemDetailAJAXLoader('Whole', true);
	
	var txtItemDetailsQty = document.getElementById('txtItemDetailsQty');
	var txtItemDetailsSaleType = document.getElementById('txtItemDetailsSaleType');
	var txtItemDetailsDisc = document.getElementById('txtItemDetailsDisc');
	var txtItemDetailsPrice = document.getElementById('txtItemDetailsPrice');
	var txtItemDetailsTotal = document.getElementById('txtItemDetailsTotal');
	var ItemDetailsInvErr = document.getElementById('ItemDetailsInvErr');
	var itemDetailsWhs = document.getElementById('itemDetailsWhs');
	
	var discPrcnt = txtItemDetailsDisc ? txtItemDetailsDisc.value.replace('%', '') : '';
	
	PrevSaleType = itemDetailsSaleType;
	itemDetailsSaleType = parseInt(txtItemDetailsSaleType.value);
	
	if (id == 3 || id == 4)
	{
		itemDetailsManPrc = true;
	}

	var ManPrc = itemDetailsManPrc ? 'Y' : 'N';

	doItemChangeField(id, false, itemDetailsID, txtItemDetailsQty, itemDetailsTreeType, itemDetailsSaleType, PrevSaleType, txtItemDetailsDisc, txtItemDetailsPrice, txtItemDetailsTotal, ItemDetailsInvErr, itemDetailsWhs, discPrcnt, ManPrc, itemDetailsCur, null);
}

function compChangeField(id, i)
{
	showItemDetailAJAXLoader('Whole', true);

	var itemCode = document.getElementById('compItem' + i);
	var compQty = document.getElementById('compQty' + i);
	var compTreeType = document.getElementById('compTreeType' + i);
	var compUnit = document.getElementById('compUnit' + i);
	var compPrevUnit = document.getElementById('compPrevUnit' + i);
	var compWhs = document.getElementById('compWhs' + i);
	var compDiscount = document.getElementById('compDiscount' + i);
	var compPrice = document.getElementById('compPrice' + i);
	var compLineTotal = document.getElementById('compLineTotal' + i);
	var compManPrc = document.getElementById('compManPrc' + i);
	var compCur = document.getElementById('compCur' + i);
	
	var compInvErr = document.getElementById('compInvErr' + i);
	
	var discPrcnt = compDiscount ? compDiscount.value.replace('%', '') : '';
	
	if (id == 3 || id == 4)
	{
		compManPrc.value = 'Y';
	}
	if (id == 1) { if (parseInt(compQty.value) <= 0) compQty.value = 1; } 
	var PrevUnit = compPrevUnit.value;
	compPrevUnit.value = compUnit.value;
	doItemChangeField(id, true, itemCode.value, compQty, compTreeType.value, compUnit.value, PrevUnit, compDiscount, compPrice, compLineTotal, compInvErr, compWhs, discPrcnt, compManPrc.value, compCur.value, i);
}

function showItemDetailAJAXLoader(id, show)
{
	var strShow = show ? '' : 'none';
	document.getElementById('tblItemDetails' + id).style.display = strShow;
	document.getElementById('tblItemDetails' + id + 'Img').style.display = strShow;
}

function closeItemDetails()
{
	jQuery('#dvItemDetails').dialog('close');
}

  jQuery(document).ready(function() {
    jQuery("#dvItemDetails").dialog({
      bgiframe: true, autoOpen: false, width: 950, height: 566, minWidth: 950, minHeight: 566, modal: true, resizable: false
    });
  });


function openItemDetails(item)
{
	var dvPopupContent = document.getElementById('dvItemDetailsContent');
		
	if (dvPopupContent.innerHTML == '') LoadItemDetails(dvPopupContent, item);
	else LoadItemDetailsData(item);

	jQuery('#dvItemDetails').dialog('open');
}

var tblVolRep = null;
var tblVolRepTimer = null;
var tblVolRepAddTop = 0;
var tblVolRepAddLeft = 0;
function clearVolRep()
{
	tblVolRepTimer = setTimeout('hideVolRep();', 100);
}

function hideVolRep()
{
	if (tblVolRep != null) tblVolRep.style.display = 'none';
}

function cancelVolHide()
{
	clearTimeout(tblVolRepTimer);
}

function CreateCartVolRep(outer, id)
{
	var strTblVolRep = '';
	strTblVolRep += '<table border="0" width="200" bgcolor="white" onmouseover="javascript:cancelVolHide();" onmouseout="javascript:clearVolRep();" id="' + id + '" cellpadding="0" style="border-style: solid;border-width: 1px; position: absolute; display: none;" cellspacing="0">' ;
	strTblVolRep += '	<caption class="FirmTlt">' + txtVolDiscount + '</caption>' ;
	strTblVolRep += '	<tr>' ;
	strTblVolRep += '		<td width="50%" class="FirmTlt3" style="text-align: center; border-bottom-style: solid; border-bottom-width: 1px; padding-bottom: 4px; border-right-style: solid; border-right-width: 1px; padding-right: 4px">' + lblItemDetailsQty + '</td>' ;
	strTblVolRep += '		<td width="50%" class="FirmTlt3" style="text-align: center; border-bottom-style: solid; border-bottom-width: 1px; padding-bottom: 4px; border-left-style: solid; border-left-width: 1px; padding-left: 4px">' + lblItemDetailsPrice + '</td>' ;
	strTblVolRep += '	</tr>' ;
	strTblVolRep += '</table>' ;
	if (!outer)
	{
		$('#dvItemDetailsContent').append(strTblVolRep);
	}
	else
	{
		$('#printThis').append(strTblVolRep);
	}
}

function displayVolRep(item, img, e, volUnit, volDate)
{
	cancelVolHide();
	$.post('Fetch/itemDetailsFetch.asp?d=' + (new Date()).toString(), { DataType: 'VD', Item: item, SaleType: volUnit, Date: volDate }, function(data)
	{
		var arrData = data.split('{S}');
		
		

		for (var i = tblVolRep.rows.length-1;i>=1;i--)tblVolRep.deleteRow(i);
		
		for (var i = 0;i<arrData.length;i++)
		{
			var strItemVolRepData = arrData[i].split('{C}');
			
			tblVolRepAddTop -= 13;			
			
			var newRow = tblVolRep.insertRow();
			
			var newCell = newRow.insertCell();
			newCell.innerHTML = strItemVolRepData[0] ;
			newCell.style.borderRightStyle = 'solid';
			newCell.style.borderRightWidth = '1px';
			newCell.style.paddingRight = '4px';
			newCell.style.textAlign = 'center'
			newCell.style.width = '50%';
			newCell.className = 'FirmTbl';
			
			newCell = newRow.insertCell();
			newCell.innerHTML = strItemVolRepData[1] + '&nbsp;';
			newCell.style.borderLeftStyle = 'solid';
			newCell.style.borderLeftWidth = '1px';
			newCell.style.paddingLeft = '4px';
			newCell.style.textAlign = 'right'
			newCell.style.width = '50%';
			newCell.className = 'FirmTbl';
		}
		
		
		var imgOffset = jQuery(img).offset();
		tblVolRep.style.left = imgOffset.left+(rtl != '' ? 23 : -200)+tblVolRepAddLeft;
		tblVolRep.style.top = imgOffset.top + tblVolRepAddTop;
		tblVolRep.style.display = '';
	});
}

function doShowItemVolRep(item, date, unit, img, e)
{
	if (document.getElementById('tblItemVolRep'))
	{
		tblVolRep = document.getElementById('tblItemVolRep');
	}
	else
	{
		CreateCartVolRep(false, 'tblItemVolRep');
		tblVolRep = document.getElementById('tblItemVolRep');
	}
	
	var dialogOffset = jQuery('#dvItemDetails').dialog().offset();

	tblVolRepAddTop = (dialogOffset.top * -1) - 6;
	tblVolRepAddLeft = (dialogOffset.left * -1) - 6;
	displayVolRep(item, img, e, unit, date);

}

function showItemVolRep(img, e)
{
	var docDate = '06/11/10';
	doShowItemVolRep(itemDetailsID, docDate, itemDetailsSaleType, img, e);

}
function showCompVolRep(img, e, i)
{
	var item = document.getElementById('compItem' + i).value;
	var unit = document.getElementById('compUnit' + i).value;

	var docDate = '06/11/10';
	doShowItemVolRep(item, docDate, unit, img, e);
}

function goItemDetTab(tab)
{
	$('#itemDetTabs').tabs('select', tab);
}

