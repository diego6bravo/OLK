var closeSmallSearch = false;
function valSmallSearch()
{
	if (document.frmSmallSearch.document.value == 'B')
		if (document.frmSmallSearch.string.value == '')
		{
			alert(txtValEnterValue);
			document.frmSmallSearch.string.focus();
			return false;
		}
	return true;
}
function goSearchCart()
{	
	document.frmSmallSearch.action = 'cart.asp';
	document.frmSmallSearch.cmd.value = 'cart';
	selViewType('B', viewTypeCount-1);
	document.getElementById('tdViewTypeCart').style.borderColor='white';
	document.frmSmallSearch.string.focus();
}
function focusSmallSearch(focus)
{
	if (focus)
	{
		closeSmallSearch = false;
		document.getElementById('trSmallSearch').style.display = '';
	}
	else
	{
		closeSmallSearch = true;
		setTimeout('doCloseSmallSearch()', 500);
	}
}
function doCloseSmallSearch()
{
	if (closeSmallSearch)
	{
		document.getElementById('trSmallSearch').style.display = 'none';
		closeSmallSearch = false;
	}
}

function valFastAdd(frm)
{
	if (frm.Item.value == '') {
		alert(txtAddItmVal);
		frm.Item.focus();
		return false;
	}
	return true;
}
var closeSmallAddItm = false;
function focusSmallAddItm(focus)
{
	if (focus)
	{
		closeSmallAddItm = false;
		for (var i = 0;i<trSmallAddItm.length;i++)
		{
			trSmallAddItm[i].style.display = '';
		}
	}
	else
	{
		closeSmallAddItm = true;
		setTimeout('doCloseSmallAddItm()', 500);
	}
}
function doCloseSmallAddItm()
{
	if (closeSmallAddItm)
	{
		for (var i = 0;i<trSmallAddItm.length;i++)
		{
			trSmallAddItm[i].style.display = 'none';
		}
		closeSmallAddItm = false;
	}
}
function chkExecCarInvAdd(e)
{
	if (e.keyCode == 13) cartInvAddItem();
}
function cartInvAddItem()
{
	if (valFastAdd(document.frmFastAdd))
	{
		var item = document.frmFastAdd.Item.value;
		var qty = document.frmFastAdd.T1.value;
		var unit = document.frmFastAdd.SaleType.value;
		var price = document.getElementById('searchCartIncPrice') ? document.frmFastAdd.precio.value : '';
		
		setFlowAlertVars('D2', (item + '{S}' + qty + '{S}' + unit + '{S}' + price + '{S}'), 'document.frmFastAdd.DocConf.value=typeIDs;document.frmFastAdd.submit();', '');
		doFlowAlert();
	}
}
