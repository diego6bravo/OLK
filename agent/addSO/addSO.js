$('#itemDetTabs').tabs();
$('#itemDetTabs').tabs("select", 0);

var updFldLineID;
var updFldCode;
var updFldDesc;
var updFldType;

function setActFlow(logNum)
{
	setFlowAlertVars('C3', '', 'document.frmAddSO.DocConf.value = typeIDs;document.frmAddSO.submit();', '');
}

function fetchValue(dataType, lineID, fldCode, fldDesc) 
{
	updFldType = dataType;
	updFldLineID = lineID;
	updFldCode = fldCode;
	updFldDesc = fldDesc;
	
	var fldID = '';
	var fldSearchStr = '';
	var DocType = '';
	var DocNum = '';
	switch (dataType)
	{
		case 0:
			fldID = 'Territory';
			fldSearchStr = fldDesc.value;
			break;
		case 1:
			fldID = 'Crd';
			fldSearchStr = fldCode.value;
			break;
		case 2:
			fldID = 'DocLink';
			fldSearchStr = '';
			DocNum = fldCode.value;
			DocType = document.getElementById('StageObjType' + lineID).value;
			break;
	}
	
	if (fldSearchStr == '' && DocNum == '') 
	{ 
		fldCode.value = ''; 
		doProcDataType();
		return; 
	} 
	
	
	if (fldSearchStr.indexOf('*') == -1 && DocNum.indexOf('*') == -1) 
	{
		$.post('topGetValueFetch.asp', { Type: fldID, searchStr: fldSearchStr, PassDesc: 'Y', DocType: DocType, DocNum: DocNum }, function(data)
			{
				if (data != '{NoData}')
				{
					var arrValues = data.split('{S}');
					updFldCode.value = dataType == 2 ? arrValues[1] : arrValues[0];
					if (updFldDesc != null) updFldDesc.value = arrValues[1];
					doProcDataType();
				}
				else
				{
					alert(txtNoData);
					updFldCode.value = '';
					if (updFldDesc != null) updFldDesc.value = '';
					doProcDataType();
				}
			});
	}
	else 
	{ 
		updFldCode.value = '';
		if (updFldDesc != null) updFldDesc.value = '';

		$.post('topGetValueSelectFetch.asp', { Type: fldID, searchStr: fldSearchStr, PassDesc: 'Y', DocType: DocType, DocNum: DocNum }, function(data)
			{ 
				if (data != '{NoData}')
				{

					$('#tblValueSelectData').empty();
					
					var txtTitle = document.getElementById('dvValueSelectTitle');
					
					var fldSelect;
					var fldDesc;
					var startCol;
					
					//Column Title
					var strRow = '<tr class="GeneralTblBold2">';
					
					switch (dataType)
					{
						case 0:
							strRow += '<td>' + txtTerritory + '</td><td>' + txtParent + '</td>';
							fldSelect = 0;
							fldDesc = 1;
							startCol = 1;
							txtTitle.innerText = txtTerritory;
							break;
						case 1:
							strRow += '<td>' + txtCode + '</td><td>' + txtName + '</td>';
							fldSelect = 0;
							fldDesc = 1;
							startCol = 0;
							txtTitle.innerText = txtBP;
							break;
						case 2:
							strRow += '<td>#</td><td>' + txtDate + '</td><td>' + txtBP + '</td><td>' + txtComments + '</td><td>' + txtDueDate + '</td>';
							fldSelect = 1;
							fldDesc = 1;
							startCol = 1;
							txtTitle.innerText = txtDocSel;
							break;
					}
					strRow += '</tr>';
					
					$('#tblValueSelectData').append(strRow);

					
					
					var arrData = data.split('{S}');
					for (var i = 0;i<arrData.length;i++)
					{
						var arrCols = arrData[i].split('{C}');
						strRow = '<tr class="GeneralTbl" style="cursor: pointer;" onmouseover="this.style.backgroundColor=\'#CDE3FC\';" onmouseout="this.style.backgroundColor=\'\';" onclick="setGetValueSelect(\'' + arrCols[fldSelect].replace('\'', '\\').replace('"', '""') + '\', \'' + arrCols[fldDesc].replace('\'', '\\').replace('"', '""') + '\');">';
						for (var j = startCol;j<arrCols.length;j++)
						{
							strRow += '<td>' + arrCols[j] + '</td>';
						}
						strRow += '</tr>';
						$('#tblValueSelectData').append(strRow);
					}
					
			
					jQuery('#dvValueSelect').dialog('open');
				}
				else
				{
					alert(txtNoData);
					updFldCode.value = '';
					if (updFldDesc != null) updFldDesc.value = '';
					doProcDataType();
				}
			});
	}
}

function setGetValueSelect(value, desc)
{
	updFldCode.value = value;
	if (updFldDesc != null) updFldDesc.value = desc;
	doProcDataType();
	
	jQuery('#dvValueSelect').dialog('close');
}

function doProcDataType()
{
	switch (updFldType)
	{
		case 0:
			doProc('Territory', 'N', updFldCode.value);
			break;
		case 1:
			doProc('ChnCrdCode', 'S', updFldCode.value);
			doProc('ChnCrdName', 'S', updFldDesc.value);
			break;
		case 2:
			doProcLine(1, updFldLineID, 'DocNumber', 'I', updFldCode.value);
			if (updFldCode.value != '')
				if (confirm(txtConfUpdAmtGrs))
				{
					$.post("addSO/addSOProcessSetAmtGrs.asp?d=" + (new Date()).toString(), { Line: updFldLineID },
					   function(data){
						var arrData = data.split('{S}');
						document.getElementById('MaxSumLoc').value = arrData[0];
						document.getElementById('WtSumLoc').value = arrData[1];
						document.getElementById('PrcntProf').value = arrData[2];
						document.getElementById('SumProfL').value = arrData[3];
						
						
						var stageNum = document.getElementById('MaxStageNum').value;
						document.getElementById('StageMaxSumLoc' + stageNum).value = arrData[0];
						document.getElementById('StageWtSumLoc' + stageNum).value = arrData[1];
						});
				}
			break;
	}
}

var isUpdating = false;
function doProc(fld, fldType, value)
{
	isUpdating = true;
	$.post("addSO/addSOProcess.asp?d=" + (new Date()).toString(), { Field: fld, FieldType: fldType, Value: value },
	   function(data){
	   	switch (fld)
	   	{
	   		case 'ChnCrdCode':
		   		var cmb = document.getElementById('ChnCrdCon');
		   		for (var i = cmb.length-1;i>0;i--)
		   			cmb.options.remove(i);
		   		
		   		var arrData = data.split('{S}');
		   		var cntData = arrData[1].split('{V}');
		   		for (var i = 0;i<cntData.length-1;i++)
		   		{
		   			var arrCnt = cntData[i].split('{C}');
		   			cmb.options[i+1] = new Option(arrCnt[1], arrCnt[0]);
		   		}
		   		cmb.value = arrData[0];
	   			break;
	   		default:
			     if (data != 'ok')
			     {
			     	alert(txtErrSaveData);
			     }
			     break;
		}
		isUpdating = false;
   });
}

function doChStageObjType(line, value)
{
	var txt = document.getElementById('StageDocNumber' + line);
	txt.value = '';
	if (value == '')
	{
		txt.className = 'inputDis';
		txt.readOnly = true;
	}
	else
	{
		txt.className = 'input';
		txt.readOnly = false;
	}

	doProcLine(1, line, 'ObjType', 'I', value);
	doProcLine(1, line, 'DocNumber', 'I', '');
}
function doProcLine(dataType, line, fld, fldType, value)
{
	$.post("addSO/addSOProcessLine.asp?d=" + (new Date()).toString(), { DataType: dataType, Line: line, Field: fld, FieldType: fldType, Value: value },
	   function(data){
	   	if (data == 'ok')
	   	{
	   		if (value == '')
	   		{
	   			clearTableRow(dataType, line);
		   	}
	   	}
	   	else
		{
			alert(txtErrSaveData);
		}
   });
}

function clearTableRow(dataType, line)
{
	var startNum;
	var tableID;
	var rowID;
	var controlID;
	switch (dataType)
	{
		case 4:
			startNum = 2;
			tableID = 'tblIntRange';
			rowID = 'intRangeNum';
			controlID = 'intRange';
			break;
		case 2:
			startNum = 1;
			tableID = 'tblBP';
			rowID = 'bpNum';
			controlID = 'NewBP';
			break;
		case 3:
			startNum = 1;
			tableID = 'tblComp';
			rowID = 'compNum';
			controlID = 'CompNew';
			break;
		case 5:
			startNum = 1;
			tableID = 'tblReason';
			rowID = 'reasonNum';
			controlID = 'ReasonNew';
			break;
	}
	var tbl = document.getElementById(tableID);
	
	tbl.deleteRow(document.getElementById(rowID + line).rowIndex);
	
	for (var i = startNum;i<tbl.rows.length-1;i++)
	{
		tbl.rows[i].cells[0].innerText = i-startNum+1;
	}
	
	document.getElementById(controlID + 'Count').value = parseInt(document.getElementById(controlID + 'Count').value)-1;
	document.getElementById(controlID + 'CountText').innerText = document.getElementById(controlID + 'Count').value;
}

function doProcStepID(line, value)
{
	$.post("addSO/addSOProcessLineFetch.asp?d=" + (new Date()).toString(), { DataType: 1, Line: line, Field: 'Step_Id', FieldType: 'I', Value: value },
	   function(data){
	   	var arrData = data.split('{S}');
	   	document.getElementById('StageClosePrcnt' + line).value = arrData[0];
	   	document.getElementById('StageMaxSumLoc' + line).value = arrData[1];
	   	document.getElementById('StageWtSumLoc' + line).value = arrData[2];
   });
}

function doProcStepClosePer(line, value)
{
	$.post("addSO/addSOProcessLineFetch.asp?d=" + (new Date()).toString(), { DataType: 1, Line: line, Field: 'ClosePrcnt', FieldType: 'N', Value: value },
	   function(data){
	   	var arrData = data.split('{S}');
	   	document.getElementById('StageClosePrcnt' + line).value = arrData[0];
	   	document.getElementById('StageWtSumLoc' + line).value = arrData[1];
   });
}

function doProcStepMaxSum(line, value)
{
	$.post("addSO/addSOProcessLineFetch.asp?d=" + (new Date()).toString(), { DataType: 1, Line: line, Field: 'MaxSumLoc', FieldType: 'N', Value: value },
	   function(data){
	   	var arrData = data.split('{S}');
	   	document.getElementById('StageMaxSumLoc' + line).value = arrData[0];
	   	document.getElementById('StageWtSumLoc' + line).value = arrData[1];
	   	document.getElementById('MaxSumLoc').value = arrData[0];
	   	document.getElementById('WtSumLoc').value = arrData[1];
	   	document.getElementById('SumProfL').value = arrData[2];

   });
}

function doProcStepWtSum(line, value)
{
	$.post("addSO/addSOProcessLineFetch.asp?d=" + (new Date()).toString(), { DataType: 1, Line: line, Field: 'WtSumLoc', FieldType: 'N', Value: value },
	   function(data){
	   	var arrData = data.split('{S}');
	   	document.getElementById('StageMaxSumLoc' + line).value = arrData[0];
	   	document.getElementById('StageWtSumLoc' + line).value = arrData[1];
	   	document.getElementById('MaxSumLoc').value = arrData[0];
	   	document.getElementById('WtSumLoc').value = arrData[1];
	   	document.getElementById('SumProfL').value = arrData[2];
   });
}

function doProcBP(line, value)
{
	$.post("addSO/addSOProcessLineFetch.asp?d=" + (new Date()).toString(), { DataType: 2, Line: line, Field: 'ParterId', FieldType: 'I', Value: value },
	   function(data){
	   	if (value != '')
	   	{
		   	var arrData = data.split('{S}');
		   	document.getElementById('BPOrlCode' + line).value = arrData[0];
		   	document.getElementById('RelatCard' + line).value = arrData[1];
		   	document.getElementById('BPMemo' + line).value = arrData[2];
		}
		else
	   		clearTableRow(2, line);
   });
}

function doNewInt(value)
{
	$.post("addSO/addSOProcessLineFetch.asp?d=" + (new Date()).toString(), { DataType: 4, Line: -1, Field: 'NewInt', FieldType: 'I', Value: value },
	   function(data){
	  	var intCount = document.getElementById('intRangeCount');

		var strLine;
		strLine = '<tr class="GeneralTbl" id="intRangeNum' + data + '">' ;
		strLine += '	<td style="width: 40px; text-align: right;">' + intCount.value + '</td>' ;
		strLine += '	<td>' ;
		strLine += '	<select class="input" size="1" name="IntRange' + data + '" style="width: 100%;" onchange="doProcLine(4, ' + data + ', \'IntId\', \'N\', this.value);"><option></option>' ;
		
		var cmbOpt = document.getElementById('IntRangeNew');
		for (var i = 1;i<cmbOpt.options.length;i++)
		{
			strLine += '	<option ' + (parseInt(value) == parseInt(cmbOpt.options[i].value) ? 'selected' : '') + ' value="' + cmbOpt.options[i].value + '">' + cmbOpt.options[i].text + '</option>' ;
		}
		strLine += '    </select></td>' ;
		strLine += '	<td style="width: 100px; text-align:center;"><input onclick="doProcLine(4, ' + data + ', \'Prmry\', \'S\', \'Y\');" type="radio" name="IntRangePrim" class="noborder" value="' + data + '"></td>' ;
		strLine += '</tr>' ;
		
		$('#tblIntRange tr:last').before(strLine)
				
		intCount.value = parseInt(intCount.value)+1;
		
		document.getElementById('intRangeCountText').innerText = intCount.value;
		
		cmbOpt.selectedIndex = 0;
   });
}

function doNewReason(value)
{
	$.post("addSO/addSOProcessLineFetch.asp?d=" + (new Date()).toString(), { DataType: 4, Line: -1, Field: 'NewReason', FieldType: 'I', Value: value },
	   function(data){
	  	var intCount = document.getElementById('ReasonNewCount');

		var strLine;
		strLine = '<tr class="GeneralTbl" id="reasonNum' + data + '">' ;
		strLine += '	<td style="width: 40px; text-align: right;">' + intCount.value + '</td>' ;
		strLine += '	<td>' ;
		strLine += '	<select class="input" size="1" name="Reason' + data + '" style="width: 100%;" onchange="doProcLine(5, ' + data + ', \'ReasonId\', \'N\', this.value);"><option></option>' ;
		
		var cmbOpt = document.getElementById('ReasonNew');
		for (var i = 1;i<cmbOpt.options.length;i++)
		{
			strLine += '	<option ' + (parseInt(value) == parseInt(cmbOpt.options[i].value) ? 'selected' : '') + ' value="' + cmbOpt.options[i].value + '">' + cmbOpt.options[i].text + '</option>' ;
		}
		strLine += '    </select></td>' ;
		strLine += '</tr>' ;

		$('#tblReason tr:last').before(strLine)
				
		intCount.value = parseInt(intCount.value)+1;
		
		document.getElementById('ReasonNewCountText').innerText = intCount.value;
		
		cmbOpt.selectedIndex = 0;
   });
}


function doNewStage(value)
{
	$.post("addSO/addSOProcessLineFetch.asp?d=" + (new Date()).toString(), { DataType: 4, Line: -1, Field: 'NewStage', FieldType: 'I', Value: value },
	   function(data){
	   	var arrData = data.split('{S}');
	   	
	  	var intCount = document.getElementById('StageCount');

		var strLine;
		
		strLine = '<tr class="GeneralTbl">' ;
		strLine += '<td style="width: 40px; text-align: right;">' + intCount.value + '</td>' ;
		strLine += '<td>' ;
		strLine += '<select class="input"size="1" name=\'StageStepID' + arrData[0] + '\' onchange="doProcStepID(' + arrData[0] + ', this.value);">' ;
		
		var cmbOpt = document.getElementById('NewStageStepID');
		for (var i = 1;i<cmbOpt.options.length;i++)
		{
			strLine += '<option ' + (parseInt(value) == parseInt(cmbOpt.options[i].value) ? 'selected' : '') + ' value="' + cmbOpt.options[i].value + '">' + cmbOpt.options[i].text + '</option>' ;
		}

		strLine += '</select></td>';
		strLine += '<td>';
		strLine += '<table cellpadding="0" cellspacing="0" border="0" width="100%">';
		strLine += '<tr class="GeneralTblBold2">' ;
		strLine += '<td align="right" width="16"><img border="0" src="images/cal.gif" id="btnStageOpenDate' + arrData[0] + '"></td>' ;
		strLine += '<td><input class="input"type="text" onclick="btnStageOpenDate' + arrData[0] + '.click();" name="StageOpenDate' + arrData[0] + '" size="12" value=\'' + arrData[5] + '\' readonly onchange="doProcLine(1, ' + arrData[0] + ', \'OpenDate\', \'D\', this.value);"></td>' ;
		strLine += '</tr>' ;
		strLine += '</table>' ;
		strLine += '</td>' ;
		strLine += '<td>' ;
		strLine += '<table cellpadding="0" cellspacing="0" border="0" width="100%">' ;
		strLine += '<tr class="GeneralTblBold2">' ;
		strLine += '<td align="right" width="16"><img border="0" src="images/cal.gif" id="btnStageCloseDate' + arrData[0] + '"></td>' ;
		strLine += '<td><input class="input" type="text" onclick="btnStageCloseDate' + arrData[0] + '.click();" name="StageCloseDate' + arrData[0] + '" size="12" value=\'' + arrData[4] + '\' readonly onchange="doProcLine(1, ' + arrData[0] + ', \'CloseDate\', \'D\', this.value);"></td>' ;
		strLine += '</tr>' ;
		strLine += '</table>' ;
		strLine += '</td>' ;
		strLine += '<td>' ;
		strLine += '<select class="input"size="1" name="StageSlpCode' + arrData[0] + '" onchange="doProcLine(1, ' + arrData[0] + ', \'SlpCode\', \'I\', this.value);">' ;
		
		var cmbSlp = document.getElementById('NewStageSlpCode');
		for (var i = 1;i<cmbSlp.options.length;i++)
		{
			strLine += '<option ' + (parseInt(arrData[6]) == parseInt(cmbSlp.options[i].value) ? 'selected' : '') + ' value="' + cmbSlp.options[i].value + '">' + cmbSlp.options[i].text + '</option>' ;
		}

		strLine += '</select></td>' ;
		strLine += '<td>' ;
		strLine += '<input onkeydown="return valKeyNumDec(event);" class="input" onfocus="this.select();"type="text" id=\'StageClosePrcnt' + arrData[0] + '\' size="8" value=\'' + arrData[1] + '\' style="text-align: right;" onchange="doProcStepClosePer(' + arrData[0] + ', this.value);"></td>' ;
		strLine += '<td>' ;
		strLine += '<input onkeydown="return valKeyNumDec(event);" class="input" onfocus="this.select();"type="text" id="StageMaxSumLoc' + arrData[0] + '" size="20" value=\'' + arrData[2] + '\' style="font-size: 8pt; text-align: right;" onchange="doProcStepMaxSum(' + arrData[0] + ', this.value);"></td>' ;
		strLine += '<td>' ;
		strLine += '<input onkeydown="return valKeyNumDec(event);" class="input" onfocus="this.select();"type="text" id=\'StageWtSumLoc' + arrData[0] + '\' size="20" value=\'' + arrData[3] + '\' style="font-size: 8pt; text-align: right;" onchange="doProcStepWtSum(' + arrData[0] + ', this.value);"></td>' ;
		strLine += '<td>' ;
		strLine += '<select class="input" size="1" name=\'StageObjType' + arrData[0] + '\' onchange="doProcLine(1, ' + arrData[0] + ', \'ObjType\', \'I\', this.value);"><option></option>' ;
		
		var cmbObjTyp = document.getElementById('NewStageObjType');
		for (var i = 1;i<cmbObjTyp.options.length;i++)
		{
			strLine += '<option value="' + cmbObjTyp.options[i].value + '">' + cmbObjTyp.options[i].text + '</option>' ;
		}

		strLine += '<option></option>' ;
		strLine += '</select></td>' ;
		strLine += '<td>' ;
		strLine += '<input onkeydown="return valKeyNum(event);" class="input" onfocus="this.select();" type="text" name=\'StageDocNumber' + arrData[0] + '\' size="20" value=\'\' onchange="doProcLine(1, ' + arrData[0] + ', \'DocNumber\', \'I\', this.value);" style="font-size: 8pt; text-align: right;"></td>' ;
		strLine += '<td class="style1">' ;
		strLine += '<select class="input" size="1" name="StageOwner' + arrData[0] + '" onchange="doProcLine(1, ' + arrData[0] + ', \'Owner\', \'I\', this.value);">' ;
		
		var cmbOwner = document.getElementById('NewStageOwner');
		for (var i = 1;i<cmbOwner.options.length;i++)
		{
			strLine += '<option ' + (parseInt(arrData[7]) == parseInt(cmbOwner.options[i].value) ? 'selected' : '') + ' value="' + cmbOwner.options[i].value + '">' + cmbOwner.options[i].text + '</option>' ;
		}

		strLine += '<option></option>' ;
		strLine += '</select></td>' ;
		strLine += '</tr>' ;

		$('#tblStages tr:last').before(strLine)
				
		intCount.value = parseInt(intCount.value)+1;
		
		document.getElementById('StageCountText').innerText = intCount.value;
		
		cmbOpt.selectedIndex = 0;
		
		var MaxStageNum = document.getElementById('MaxStageNum');
		var disLineNum = MaxStageNum.value;
		
		document.getElementById('StageStepID' + disLineNum).className = 'inputDis';
		document.getElementById('StageStepID' + disLineNum).disabled = true;
		document.getElementById('btnStageOpenDate' + disLineNum).disabled = true;
		document.getElementById('StageOpenDate' + disLineNum).className = 'inputDis';
		document.getElementById('btnStageCloseDate' + disLineNum).disabled = true;
		document.getElementById('StageCloseDate' + disLineNum).className = 'inputDis';
		document.getElementById('StageSlpCode' + disLineNum).className = 'inputDis';
		document.getElementById('StageSlpCode' + disLineNum).disabled = true;
		document.getElementById('StageClosePrcnt' + disLineNum).className = 'inputDis';
		document.getElementById('StageClosePrcnt' + disLineNum).readonly = true;
		document.getElementById('StageMaxSumLoc' + disLineNum).className = 'inputDis';
		document.getElementById('StageMaxSumLoc' + disLineNum).readonly = true;
		document.getElementById('StageWtSumLoc' + disLineNum).className = 'inputDis';
		document.getElementById('StageWtSumLoc' + disLineNum).readonly = true;
		document.getElementById('StageObjType' + disLineNum).className = 'inputDis';
		document.getElementById('StageObjType' + disLineNum).disabled = true;
		document.getElementById('StageDocNumber' + disLineNum).className = 'inputDis';
		document.getElementById('StageDocNumber' + disLineNum).readonly = true;
		document.getElementById('StageOwner' + disLineNum).className = 'inputDis';
		document.getElementById('StageOwner' + disLineNum).disabled = true;
		
		MaxStageNum.value = arrData[0];
		
		Calendar.setup({
		    inputField     :    "StageOpenDate" + arrData[0],     // id of the input field
		    ifFormat       :    CalendarFormat,      // format of the input field
		    button         :    "btnStageOpenDate" + arrData[0],  // trigger for the calendar (button ID)
		    align          :    "Bl",           // alignment (defaults to "Bl")
		    singleClick    :    true
		});
		Calendar.setup({
		    inputField     :    "StageCloseDate" + arrData[0],     // id of the input field
		    ifFormat       :    CalendarFormat,      // format of the input field
		    button         :    "btnStageCloseDate" + arrData[0],  // trigger for the calendar (button ID)
		    align          :    "Bl",           // alignment (defaults to "Bl")
		    singleClick    :    true
		});

   });
}


function doNewBP(value)
{
	$.post("addSO/addSOProcessLineFetch.asp?d=" + (new Date()).toString(), { DataType: 4, Line: -1, Field: 'NewBP', FieldType: 'I', Value: value },
	   function(data){
		var arrData = data.split('{S}');
		
	  	var intCount = document.getElementById('NewBPCount');

		var strLine;
		strLine = '<tr class="GeneralTbl" id="bpNum' + arrData + '">' ;
		strLine += '<td style="width: 40px; text-align: right; height: 24px;">' + intCount.value + '</td>' ;
		strLine += '<td style="height: 24px">' ;
		strLine += '<select class="input" size="1" name="BPPartnerId' + arrData + '" style="width: 222px; font-size:10px; font-family:Verdana;" onchange="doProcBP(' + arrData + ', this.value);">' ;
		strLine += '<option></option>' ;
		
		var cmbOpt = document.getElementById('NewBP');
		for (var i = 1;i<cmbOpt.options.length;i++)
		{
			strLine += '	<option ' + (parseInt(value) == parseInt(cmbOpt.options[i].value) ? 'selected' : '') + ' value="' + cmbOpt.options[i].value + '">' + cmbOpt.options[i].text + '</option>' ;
		}
		
		strLine += '</select></td>' ;
		strLine += '<td style="height: 24px">' ;
		strLine += '<select class="input" size="1" id="BPOrlCode' + arrData + '" style="width: 222px; font-size:10px; font-family:Verdana;">' ;
		
		var cmbNewBPOrlCode = document.getElementById('NewBPOrlCode');
		for (var i = 1;i<cmbOpt.options.length;i++)
		{
			strLine += '	<option ' + (parseInt(arrData[1]) == parseInt(cmbNewBPOrlCode.options[i].value) ? 'selected' : '') + ' value="' + cmbNewBPOrlCode.options[i].value + '">' + cmbNewBPOrlCode.options[i].text + '</option>' ;
		}
		
		strLine += '</select></td>' ;
		strLine += '<td style="width: 120px; height: 24px;"><input type="text" class="inputDis" id="RelatCard' + arrData + '" maxlength="15" value="' + arrData[2] + '" readonly style="width: 100%;"></td>' ;
		strLine += '<td style="height: 24px"><input type="text" class="input" maxlength="50" value="' + arrData[3] + '" id="BPMemo' + arrData + '" style="width: 100%;" onchange="doProcLine(2, ' + arrData + ', \'Memo\', \'S\', this.value);"></td>' ;
		strLine += '</tr>' ;
		
		$('#tblBP tr:last').before(strLine)
				
		intCount.value = parseInt(intCount.value)+1;
		
		document.getElementById('NewBPCountText').innerText = intCount.value;
		
		cmbOpt.selectedIndex = 0;
   });
}
function doCompNew(value)
{
	$.post("addSO/addSOProcessLineFetch.asp?d=" + (new Date()).toString(), { DataType: 4, Line: -1, Field: 'NewComp', FieldType: 'I', Value: value },
	   function(data){
	   	var arrData = data.split('{S}');

	  	var intCount = document.getElementById('CompNewCount');

		var strLine;
		strLine = '<tr class="GeneralTbl" id="compNum' + arrData[0] + '">' ;
		strLine += '<td style="width: 40px; text-align: right;">' + intCount.value + '</td>' ;
		strLine += '<td>' ;
		strLine += '<select class="input" size="1" name="CompPartnerId' + arrData[0] + '" style="width: 222px; font-size:10px; font-family:Verdana;" onchange="doProcCompetition(' + arrData[0] + ', this.value);">' ;
		strLine += '<option></option>' ;

		var cmbOpt = document.getElementById('CompNew');
		for (var i = 1;i<cmbOpt.options.length;i++)
		{
			strLine += '<option ' + (parseInt(value) == parseInt(cmbOpt.options[i].value) ? 'selected' : '') + ' value="' + cmbOpt.options[i].value + '">' + cmbOpt.options[i].text + '</option>' ;
		}
		strLine += '</td><td>';
		var threatDesc = txtLow;
		switch (parseInt(arrData[2]))
		{
			case 2:
				threatDesc = txtMedium;
				break;
			case 3:
				threatDesc = txtHigh;
				break;
		}
		strLine += '<input type="text" class="inputDis" maxlength="15" id="Threat' + arrData[0] + '" value="' + threatDesc + '" readonly style="width: 100%;"></td>' ;
		strLine += '<td><input type="text" id="CompMemo' + arrData[0] + '" class="input" maxlength="50" value="' + arrData[1] + '" style="width: 100%;" onchange="doProcLine(3, ' + arrData[0] + ', \'Memo\', \'S\', this.value);"></td>' ;
		strLine += '<td style="text-align: center;">' ;
		strLine += '<input id="CompChkWon' + arrData[0] + '" type="checkbox" class="noborder" value="Y" onclick="doProcLine(3, ' + arrData[0] + ', \'Won\', \'S\', (this.checked ? \'Y\' : \'N\'));"></td>' ;
		strLine += '</tr>' ;
				
		$('#tblComp tr:last').before(strLine)
				
		intCount.value = parseInt(intCount.value)+1;
		
		document.getElementById('CompNewCountText').innerText = intCount.value;
		
		cmbOpt.selectedIndex = 0;
   });
}

function doProcCompetition(line, value)
{
	$.post("addSO/addSOProcessLineFetch.asp?d=" + (new Date()).toString(), { DataType: 3, Line: line, Field: 'CompetId', FieldType: 'I', Value: value },
	   function(data){
		if (value != '')
		{
			var arrData = data.split('{S}');
			var threatDesc = txtLow;
			switch (parseInt(arrData[1]))
			{
				case 2:
					threatDesc = txtMedium;
					break;
				case 3:
					threatDesc = txtHigh;
					break;
			}
			document.getElementById('Threat' + line).value = threatDesc;
			document.getElementById('CompMemo' + line).value = arrData[0];
			document.getElementById('CompChkWon' + line).checked = false;
		}
		else
			clearTableRow(3, line);
   });
}

Calendar.setup({
    inputField     :    "OpenDate",     // id of the input field
    ifFormat       :    CalendarFormat,      // format of the input field
    button         :    "btnOpenDate",  // trigger for the calendar (button ID)
    align          :    "Bl",           // alignment (defaults to "Bl")
    singleClick    :    true
});

Calendar.setup({
    inputField     :    "txtPredDate",     // id of the input field
    ifFormat       :    CalendarFormat,      // format of the input field
    button         :    "btnPredDate",  // trigger for the calendar (button ID)
    align          :    "Bl",           // alignment (defaults to "Bl")
    singleClick    :    true
});


function SetCloseData(fieldID, value)
{
	var openDate = document.frmAddSO.OpenDate.value;
	var difNum = document.frmAddSO.PredDateQty.value;
	var difType = document.frmAddSO.DifType.value;
	var predDate = document.frmAddSO.txtPredDate.value;
	$.post("addSO/addSOProcessCloseData.asp?d=" + (new Date()).toString(), { FieldID: fieldID, OpenDate: openDate, DifNum: difNum, DifType: difType, PredDate: predDate },
	   function(data){
	     var arrData = data.split('{S}');
	     document.frmAddSO.PredDateQty.value = arrData[0];
	     document.frmAddSO.DifType.value = arrData[1];
	     document.frmAddSO.txtPredDate.value = arrData[2];
   });
}

function SetSumData(fieldID, value)
{
	var maxSumLoc = document.frmAddSO.MaxSumLoc;
	var wtSumLoc = document.frmAddSO.WtSumLoc;
	var prcntProf = document.frmAddSO.PrcntProf;
	var sumProfL = document.frmAddSO.SumProfL;
	$.post("addSO/addSOProcessSumData.asp?d=" + (new Date()).toString(), { FieldID: fieldID, MaxSumLoc: maxSumLoc.value, WtSumLoc: wtSumLoc.value, PrcntProf: prcntProf.value, SumProfL: sumProfL.value },
	   function(data){
	     var arrData = data.split('{S}');
	     maxSumLoc.value = arrData[0];
	     wtSumLoc.value = arrData[1];
	     sumProfL.value = arrData[2];
	     prcntProf.value = arrData[3];
	     
	     var stageNum = document.getElementById('MaxStageNum').value;
	     document.getElementById('StageMaxSumLoc' + stageNum).value = arrData[0];
	     document.getElementById('StageWtSumLoc' + stageNum).value = arrData[1];

	     switch (fieldID)
	     {
	     	case 1:
	     		wtSumLoc.select();
	     		break;
	     	case 2:
	     		prcntProf.select();
	     		break;
	     	case 4:
	     		sumProfL.select();
	     		break;
	     		
	     }
	     
   });
}
  jQuery(document).ready(function() {
    jQuery("#dvValueSelect").dialog({
      bgiframe: true, autoOpen: false, width: 450, height: 566, minWidth: 450, minHeight: 566, modal: true, resizable: false
    });
  });

function valKeyNumSearch(e)
{
	switch (e.keyCode)
	{
		case 8:
		case 9:
		case 16:
		case 38:
		case 36:
		case 46:
		case 48:
		case 49:
		case 50:
		case 51:
		case 52:
		case 53:
		case 54:
		case 55:
		case 56:
		case 57:
		case 96:
		case 97:
		case 98:
		case 99:
		case 100:
		case 101:
		case 102:
		case 103:
		case 104:
		case 105:
		case 106: //*
		case 56: //*
		case 37: // Left
		case 39: //Right
			return true;
	}
	return false;
}
