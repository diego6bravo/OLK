var SaveImgField;
var SaveImgImage;
var SaveImgMaxSize;
var objFieldType;

function setItemFlow(logNum)
{
	setFlowAlertVars('A1', '', 'document.frmAddItem.DocConf.value = typeIDs;document.frmAddItem.doSubmitAdd.value=\'Y\';document.frmAddItem.submit();', 'agentItemSubmit.asp?RetVal=' + logNum + '&Confirm=Y');
}

function enableCombosVirtual(enable)
{
	var chkPrchseItem = document.getElementById('chkPrchseItem');
	var chkInvntItem = document.getElementById('chkInvntItem');
	var chkOlkCombo = document.getElementById('chkOlkCombo');
	var cmbManCostPrc = document.getElementById('cmbManCostPrc');
	var chkCmbShowComp = document.getElementById('chkCmbShowComp');
	var chkFatherShowPrice = document.getElementById('chkFatherShowPrice');
	var chkShowCompPrice = document.getElementById('chkShowCompPrice');

	if (enable)
	{
		chkOlkCombo.checked = true;
		chkOlkCombo.disabled = true;
		enableCombos(true, false);
		
		chkPrchseItem.disabled = true;
		chkPrchseItem.checked = false;
		doProc('PrchseItem', 'S', 'N');
		chkInvntItem.disabled = true;
		chkInvntItem.checked = false;
		doProc('InvntItem', 'S', 'N');
		
		cmbManCostPrc.disabled = false;
		chkCmbShowComp.disabled = false;
		chkFatherShowPrice.disabled = false;
		chkShowCompPrice.disabled = false;
		doProcShowComp(false);
	}
	else
	{
		chkPrchseItem.disabled = false;
		chkInvntItem.disabled = false;
		chkOlkCombo.disabled = false;
		LockShowComp();
	}
	doProcCmb('Virtual', 'S', GetYesNo(enable));
}

function doProcShowPrice(id, checked)
{
	doProcCmb('Show' + id + 'Price', 'S', GetYesNo(checked));
	var chk = document.getElementById('chkAllowChange' + id + 'Price');
	if (!checked)
	{
		chk.checked = false;
		chk.disabled = true;
		doProcCmb('AllowChange' + id + 'Price', 'S', GetYesNo(checked));
	}
	else
	{
		chk.disabled = false;
	}
}

function enableCombos(enable, lockShowComp)
{
	document.getElementById('ttlCmbData').style.display = enable ? '' : 'none';
	document.getElementById('trCmbData').style.display = enable ? '' : 'none';
	document.getElementById('ttlCmbComps').style.display = enable ? '' : 'none';
	document.getElementById('trCmbCompsData').style.display = enable ? '' : 'none';
	
	
	var chkSellItem = document.getElementById('chkSellItem');
	
	
	if (enable)
	{
		chkSellItem.disabled = true;
		chkSellItem.checked = true;
		doProc('SellItem', 'S', 'Y');
		if (lockShowComp)
		{
			LockShowComp();
		}
	}
	else
	{
		chkSellItem.disabled = false;
	}
	doProcCmb('EnableCombo', 'S', GetYesNo(enable));
}

function doProcShowComp(checked)
{
	doProcCmb('ShowComp', 'S', GetYesNo(checked));
	
	var chkFatherShowPrice = document.getElementById('chkFatherShowPrice');
	
	if (!checked)
	{
		chkFatherShowPrice.checked = true;
		chkFatherShowPrice.disabled = true;
		doProcShowPrice('Father', true);
	}
	else
	{
		chkFatherShowPrice.disabled = false;
	}
}

function LockShowComp()
{
	var chkCmbShowComp = document.getElementById('chkCmbShowComp');
	var chkFatherShowPrice = document.getElementById('chkFatherShowPrice');
	var chkShowCompPrice = document.getElementById('chkShowCompPrice');
	
	chkCmbShowComp.disabled = true;
	chkCmbShowComp.checked = true;
	doProcCmb('ShowComp', 'S', 'Y');
	
	chkFatherShowPrice.disabled = true;
	chkFatherShowPrice.checked = true;
	doProcShowPrice('Father', true);
	
	chkShowCompPrice.disabled = true;
	chkShowCompPrice.checked = true;
	doProcShowPrice('Comp', true);
	
	document.getElementById('cmbManCostPrc').disabled = true;
}

function getImg(Field, Img, MaxSize)
{
	SaveImgField = Field;
	SaveImgImage = Img;
	SaveImgMaxSize = MaxSize;
	Start('upload/fileupload.aspx?ID=' + dbID + '&style=../design/' + selDec + '/style/stylePopUp.css',400,100,'no')
}

function Start(page, w, h, s) 
{
	OpenWin = this.open(page, "ImagePicker", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable=no, width="+w+",height="+h);
}

function changepic(value) 
{
	SaveImgField.value = value;
	SaveImgImage.src = "pic.aspx?filename=" + value + "&MaxSize=" + SaveImgMaxSize + '&dbName=' + dbName;
	doProc(SaveImgField.name, 'S', value);
}

function datePicker(page, w, h, s, r, o, pType) 
{
	objField = o;
	objFieldType = pType;
	OpenWin = this.open(page, "datePicker", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable="+r+", width="+w+",height="+h);
	OpenWin.focus();
}

function setTimeStamp(Action, value) 
{ 
	objField.value = value;
	doProc(objField.name, objFieldType, value);
}

function goAdd(Confirm, DocConf)
{
	document.frmAddItem.Confirm.value = Confirm;
	document.frmAddItem.DocConf.value = DocConf;
	document.frmAddItem.btnAdd.click();
}

function doProc(fld, fldType, value)
{
	$.post("addItem/addItemProcess.asp?d=" + (new Date()).toString(), { Field: fld, FieldType: fldType, Value: value },
	   function(data){
	     if (data != 'ok')
	     {
	     	alert(txtErrSaveData);
	     }
   });
}

function doProcItem(item)
{
	$.post("addItem/addItemProcess.asp?d=" + (new Date()).toString(), { Field: 'ItemCode', FieldType: 'S', Value: item },
	   function(data){
	   	document.getElementById('dvCodeErr').style.display = data == 'Y' ? '' : 'none';
	   	document.frmAddItem.btnAdd.disabled = data == 'Y';
	   	document.frmAddItem.ItemCode.style.backgroundColor = data == 'Y' ? '#FFD2A6' : '';
   });
}

function doProcCmb(fld, fldType, value)
{
	$.post("addItem/addItemProcess.asp?d=" + (new Date()).toString(), { Field: fld, FieldType: fldType, Value: value, ProcType: 'Cmb' },
	   function(data){
	     if (data != 'ok')
	     {
	     	alert(txtErrSaveData);
	     }
   });
}

function doProcCmbComp(lineID, fld, fldType, value)
{
	$.post("addItem/addItemProcess.asp?d=" + (new Date()).toString(), { LineID: lineID, Field: fld, FieldType: fldType, Value: value, ProcType: 'CmbComp' },
	   function(data){
	     if (data != 'ok')
	     {
	     	alert(txtErrSaveData);
	     }
   });
}

function doProcCompLock(lineID, lockID)
{
	var lockVal = document.getElementById('hdLockVal' + lockID + lineID).value
	doProcCmbComp(lineID, 'Lock' + lockID, 'S', lockVal);
}

function doProcCompNum(lineID, fld, fieldID)
{
	switch (fieldID)
	{
		case 'DiscPrcnt':
			chkThis(fld, 'B', '%', '', 100); 
			break;
		default:
			chkThis(fld, 'I', '', '', ''); 
			break;
	}
	 if(fld.value=='')fld.value=1;
	 doProcCmbComp(lineID, fieldID, 'N', fld.value);
}

function doChkUdfVal(fieldID, chk)
{
	$.post('addItem/addItemProcess.asp?d=' + (new Date()).toString(), { ProcType: 'CheckUdfFilterValue', FieldID: fieldID, Checked: GetYesNo(chk.checked), Value: chk.value }, function(data)
	{
		if (data != 'ok')
		{
			alert('Error');
		}
	});
	
}

function showUDF(id)
{
	if (id == -1) id = '_1';
	
	var tdShowUDF = document.getElementById('tdShowUDF' + id);
	var trUDF = document.getElementById('trUDF' + id);

	trUDF.style.display = '';
	tdShowUDF.innerHTML = '[-]';
}

function doLoadField(cmb)
{
	if (cmb.selectedIndex > 0)
	{
		loadUDFValues(cmb.options[cmb.selectedIndex].text, cmb.value);
		$('#' + cmb.id + ' option[value=\'' + cmb.value + '\']').remove();
		cmb.selectedIndex = 0;
		document.getElementById('trAddField').style.display = cmb.options.length > 1 ? '' : 'none';
	}
}

function loadUDFValues(desc, fieldID)
{
	$.post('addItem/addItemProcess.asp?d=' + (new Date()).toString(), { ProcType: 'GetUDFData', FieldID: fieldID }, function(data)
	{
		var udfValues = '';
		udfValues += '<tr><td colspan="2"><hr style="height: 1px;"></td></tr>';
		udfValues += '<tr class="GeneralTbl"><td class="GeneralTblBold2" style="vertical-align: top; padding-top: 2px;">' + desc + '</td><td><div style="height: 100px; overflow: auto; overflow-x: none;"><table cellpadding="0" cellspacing="0" width="100%">';
		
		var arrData = data.split('{V}');
		for (var j = 0;j<arrData.length;j++)
		{
			var arrFldVals = arrData[j].split('{S}');
			udfValues += '<tr class="GeneralTbl"><td width="10"><input type="checkbox" class="noborder" onclick="doChkUdfVal(\'' + fieldID + '\', this);" id="chkUdfVal' + fieldID + '_' + j + '" value="' + arrFldVals[0].replace('"', '""') + '" ' + (arrFldVals[2] == 'Y' ? 'checked' : '') + '></td><td><label for="chkUdfVal' + fieldID + '_' + j + '">' + arrFldVals[1] + '</label></td></tr>';
		}
		
		udfValues += '</table></div></td></tr>' 
		$('#tbFieldData').find('tr').end().append(udfValues);
	});
}

function VerfyQuery()
{	var valFilterQuery = document.getElementById('valFilterQuery');
	var btnVerfyFilter = document.getElementById('btnVerfyFilter');
	var txtFilterQry = document.getElementById('txtFilterQry');
	
	
	$.post('addItem/addItemProcess.asp?d=' + (new Date()).toString(), { Field: 'Filter', Query: txtFilterQry.value }, function(data)
	{
		if (data != 'ok')
		{
			alert(txtError + ': ' + data.split('{S}')[1]);
		}
		else
		{
			valFilterQuery.value = 'N';
			btnVerfyFilter.src='images/btnValidateDis.gif';
			btnVerfyFilter.style.cursor = 'hand';
		}
	});
}

function SetCheckListData(chk, id, data, chkColID)
{
	var arrData = data.split('{F}');
	
	var hasData = (data != '');
	if (chk.length)
	{
		for (var i = 0;i<chk.length;i++)
		{
			SetCheckListDataCheck(chk[i], hasData, arrData, chkColID);
		}		
	}
	else
	{
		SetCheckListDataCheck(chk, hasData, arrData, chkColID);
	}
	
	document.getElementById('img' + id).src = hasData ? 'images/arrow_down.gif' : 'images/' + rtl + 'right.gif';
	document.getElementById('dv' + id).style.display = hasData ? '' : 'none';
}

function doProcFilter(filterType, chk)
{
	$.post('addItem/addItemProcess.asp?d=' + (new Date()).toString(), { ProcType: 'CheckFilterValue', FilterType: filterType, Checked: GetYesNo(chk.checked), Value: chk.value }, function(data)
	{
		if (data != 'ok')
		{
			alert('Error');
		}
	});
}

function showHideFilter(id)
{
	var dvShow = document.getElementById('dv' + id).style.display == 'none';
	
	document.getElementById('img' + id).src = dvShow ? 'images/arrow_down.gif' : 'images/' + rtl + 'right.gif';
	document.getElementById('dv' + id).style.display = dvShow ? '' : 'none';
	
	if (id == 'Qry') document.getElementById('txtOITMFilter').style.display = dvShow ? '' : 'none';
}

function SetCheckListDataCheck(chk, loadData, data, colID)
{
	var chkVal = chk.value;
	
	var found = false;
	if (loadData)
	{
		for (var i = 0;i<data.length;i++)
		{
			var arrData = data[i].split('{C}');
			if (chkVal == arrData[colID])
			{
				found = true;
				break;
			}
		}
	}
	
	chk.checked = found;
}

var updFldLineID;
var updFldCode;
var updFldDesc;
function getItem(lineID, fldCode, fldDesc) 
{
	if (fldCode.value == '') 
	{ 
		fldDesc.value = ''; 
		doProcCmbComp(lineID, 'ItemCode', 'S', '');
		return; 
	} 
	
	updFldLineID = lineID;
	updFldCode = fldCode;
	updFldDesc = fldDesc;
	
	if (fldCode.value.indexOf('*') == -1) 
	{
		$.post('topGetValueFetch.asp', { Type: 'Itm', PassDesc: 'Y', searchStr: fldCode.value }, function(data)
			{
				if (data != '{NoData}')
				{
					var arrValues = data.split('{S}');
					updFldCode.value = arrValues[0];
					updFldDesc.value = arrValues[1];
					doProcCmbComp(lineID, 'ItemCode', 'S', updFldCode.value);
				}
				else
				{
					launchSelect(updFldCode.value)
				}
			});
	}
	else 
	{ 
		launchSelect(fldCode.value); 
	}
}
function delComp(lineID)
{
	if (confirm(txtConfRemComp))
	{
		$.post("addItem/addItemComboProcess.asp?d=" + (new Date()).toString(), { CmdType: 'DelCmbComp', LineID: lineID },
		   function(data){
		    if (data == 'ok')
		    {
		    	$('#trComp' + lineID).remove();
		    }
	   });
		
	}
}
function addComp()
{
	$.post("addItem/addItemComboProcess.asp?d=" + (new Date()).toString(), { CmdType: 'AddCmbComp' },
	   function(data){
	    doAddComp(data);
   });
}
function doAddComp(data)
{
	var arrData = data.split('{S}');
	
	var optPriceList = '';
	var arrPriceList = arrData[1].split('{L}');
	for (var i = 0;i<arrPriceList.length;i++)
	{
		var arrPriceData = arrPriceList[i].split('{D}');
		optPriceList += '<option value="' + arrPriceData[0] + '">' + arrPriceData[1] + '</option>';
	}
	
	var optWhsList = '';
	var arrWhsList = arrData[2].split('{W}');
	for (var i = 0;i<arrWhsList.length;i++)
	{
		var arrWhsData = arrWhsList[i].split('{D}');
		optWhsList += '<option value="' + arrWhsData[0] + '">' + arrWhsData[1] + '</option>';
	}
	
	var lineID = arrData[0];
	
	var str = '			<tr class="GeneralTbl" id="trComp' + lineID + '"><input type="hidden" name="CompLineID" value="' + lineID + '">' +  
			'				<td>' +
			'				<table cellpadding="0" cellspacing="0">' +  
			'				<tr>' +  
			'					<td><input type="text" id="txtCompItem' + lineID + '" value="" size="20" maxlength="20" onchange="javascript:getItem(' + lineID + ', this, document.getElementById(\'txtCompDesc' + lineID + '\'));" onfocus="this.select();">' +  
			'					</td>' +  
			'					<td>' +  
			'					<img id="imgLockItm' + lineID + '" src="images/icon_unlock.jpg" onclick="ClickUnlock(\'Itm' + lineID + '\');doProcCompLock(' + lineID + ', \'Itm\');" style="cursor: pointer;">' +
			'					<input type="hidden" id="hdLockValItm' + lineID + '" name="LockItm' + lineID + '" value="N">' +
			'					</td>' +  
			'				</tr>' +  
			'				</table>' +
			'				</td>' +  
			'				<td><input type="text" id="txtCompDesc' + lineID + '" value="" size="40" readonly class="inputDes"></td>' +  
			'				<td style="text-align: center;"><input type="checkbox" id="chkCompLocked' + lineID + '" checked value="Y" style="border: solid 0px; background: background-image;" onclick="doProcCmbComp(' + lineID + ', \'Locked\', \'S\', GetYesNo(this.checked));"></td>' +  
			'				<td style="text-align: right;">' +
			'				<table cellpadding="0" cellspacing="0">' +  
			'				<tr>' +  
			'					<td><input type="text" id="txtCompQty' + lineID + '" value="1" size="5" style="text-align: right;" onchange="doProcCompNum(' + lineID + ', this, \'Quantity\');" onfocus="this.select();">' +  
			'					</td>' +  
			'					<td>' +  
			'					<img id="imgLockQty' + lineID + '" src="images/icon_unlock.jpg" onclick="ClickUnlock(\'Qty' + lineID + '\');doProcCompLock(' + lineID + ', \'Qty\');" style="cursor: pointer;">' +
			'					<input type="hidden" id="hdLockValQty' + lineID + '" name="LockQty' + lineID + '" value="N">' +
			'					</td>' +  
			'				</tr>' +  
			'				</table>' +
			'				</td>' +  
			'				<td style="text-align: right; display: none;"><input type="text" id="txtCompLines' + lineID + '" value="1" size="2" style="text-align: right;" onchange="doProcCompNum(' + lineID + ', this, \'Lines\');" onfocus="this.select();"></td>' +  
			'				<td><select id="cmbCompWhs' + lineID + '" size="1" onchange="doProcCmbComp(' + lineID + ', \'WhsCode\', \'S\', this.value);">' +  
			'				<option value="">' + txtDefault + '</option>' + optWhsList +  
			'				</select></td>' +  
			'				<td><select id="cmbCompCstPrcLst' + lineID + '" size="1" onchange="doProcCmbComp(' + lineID + ', \'AlterCostPrcList\', \'I\', this.value);">' +  
			'				<option value="">' + txtDefault + '</option>' + optPriceList +  
			'				</select></td>' +  
			'				<td><select id="cmbCompSalPrcLst' + lineID + '" size="1" onchange="doProcCmbComp(' + lineID + ', \'AlterSalePrcList\', \'I\', this.value);">' +  
			'				<option value="">' + txtDefault + '</option>' + optPriceList +  
			'				</select></td>' +  
			'				<td>' +
			'				<table cellpadding="0" cellspacing="0">' +  
			'				<tr>' +  
			'					<td><input type="text" id="txtDiscPrcnt' + lineID + '" value="0" size="5" style="text-align: right;" onchange="doProcCompNum(' + lineID + ', this, \'DiscPrcnt\');" onfocus="this.select();" onkeydown="return valKeyNum(event);">' +  
			'					</td>' +  
			'					<td>' +  
			'					<img id="imgLockDisc' + lineID + '" src="images/icon_lock.jpg" onclick="ClickUnlock(\'Disc' + lineID + '\');doProcCompLock(' + lineID + ', \'Disc\');" style="cursor: pointer;">' +
			'					<input type="hidden" id="hdLockValDisc' + lineID + '" name="LockDisc' + lineID + '" value="Y">' +
			'					</td>' +  
			'				</tr>' +  
			'				</table>' +
			'				</td>' +  
			'				<td width="16"><img id="btnDel' + lineID + '" src="ventas/images/cancel_x.gif" style="cursor: pointer;" onclick="delComp(' + lineID + ');"></td>' +  
			'			</tr>';
	$('#tbCmbData').append(str);
	
	chkThis(document.getElementById('txtDiscPrcnt' + lineID), 'B', '%', '', 100); 
}
function launchSelect(Value)
{
	var retVal = window.showModalDialog('topGetValueSelect.asp?Type=Itm&PassDesc=Y&Value=' + Value,'','dialogWidth:500px;dialogHeight:500px');
	if (retVal != '' && retVal != null)
	{
		var arrValues = retVal.split('{S}');
		updFldCode.value = arrValues[0];
		updFldDesc.value = arrValues[1];
	} 
	else 
	{ 
		updFldCode.value = '';
		updFldDesc.value = '';
	}
	doProcCmbComp(updFldLineID, 'ItemCode', 'S', updFldCode.value);
}

function GetYesNo(value)
{
	return value ? 'Y' : 'N';
}