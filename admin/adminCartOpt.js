var rsHasVars = false;

function valFrm()
{
	rowName = document.form1.Name;
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

function changeObject(obj)
{
	for (var i = document.getElementById('tblLinkVars').rows.length-1;i>0;i--)
	{
		document.getElementById('tblLinkVars').deleteRow(i);
	}
	if (obj != '')
	{
		document.getElementById('trRepVars').style.display = '';
		document.frmGetRSVars.rsIndex.value = obj;
		document.frmGetRSVars.submit();
	}
	else
	{
		document.getElementById('trRepVars').style.display = 'none';
	}
}

function setRSVars(rsVars)
{
	if (rsVars.length == 0)
	{
		rsHasVars = false;
		var newRow = tblLinkVars.insertRow(tblLinkVars.rows.length);
		var newCell = newRow.insertCell(newRow.cells.length);
		newCell.colSpan = "3";
		newCell.innerHTML = 	'	<p align="center">' +
								'	<font face="Verdana" size="1" color="#4783C5">' +
								'	' + LtxtNoRSVars + '</font>';
	}
	else
	{
		rsHasVars = true;
		for (var i = 0;i<rsVars.length;i++)
		{
			var newRow = tblLinkVars.insertRow(tblLinkVars.rows.length);
			
			var newCell = newRow.insertCell(newRow.cells.length);
			newCell.innerHTML = '<font face="Verdana" size="1" color="#4783C5">@' + rsVars[i].varVar +
								(rsVars[i].varNotNull == 'Y' ? '<font color="red">*</font>' : '') + '</font>';
				
			newCell = newRow.insertCell(newRow.cells.length);
			newCell.innerHTML = '<font face="Verdana" size="1" color="#4783C5">' +
								'<input type="radio" class="OptionButton" style="background:background-image" value="F" checked name="valBy' + rsVars[i].varVar + '" id="rdFld' + rsVars[i].varVar + '" onclick="changeValBy(\'' + rsVars[i].varVar + '\',\'F\');">' +
								'<label for="rdFld' + rsVars[i].varVar + '">' + DtxtVariable + '</label>' +
								'<input class="OptionButton" style="background:background-image" type="radio" name="valBy' + rsVars[i].varVar + '" value="V" id="rdVal' + rsVars[i].varVar + '" onclick="changeValBy(\'' + rsVars[i].varVar + '\',\'V\');">' +
								'<label for="rdVal' + rsVars[i].varVar + '">' + DtxtValue + '</label>' +
								'<input class="OptionButton" style="background:background-image" type="radio" name="valBy' + rsVars[i].varVar + '" value="Q" id="rdQry' + rsVars[i].varVar + '" onclick="changeValBy(\'' + rsVars[i].varVar + '\',\'Q\');">' +
								'<label for="rdQry' + rsVars[i].varVar + '">' + DtxtQuery + '</label>' +
								'</font>';
								
			newCell = newRow.insertCell(newRow.cells.length);
			newCell.innerHTML = '<font color="#4783C5" face="Verdana" size="1">' +
								'<table border="0" id="tblValDat' + rsVars[i].varVar + '" cellspacing="0" cellpadding="0" style="display: none">' + 
								'	<tr>' +
								'		<td><img border="0" src="images/cal.gif" id="btnValDatImg' + rsVars[i].varVar + '" width="16" height="16" style="float:left;padding-left:1px;padding-top:1px"></td>' +
								'		<td>' +
								'		<input type="text" readonly name="colValDat' + rsVars[i].varVar + '" size="12" value="" onclick="btnValDatImg' + rsVars[i].varVar + '.click()"></td>' +
								'		<td><img border="0" src="images/remove.gif" style="cursor: hand" onclick="javascript:document.form2.colValDat' + rsVars[i].varVar + '.value=\'\';"></td>' +
								'	</tr>' +
								'</table>' +
								'<input style="display: none" type="text" name="valValueV' + rsVars[i].varVar + '" id="valValueV' + rsVars[i].varVar + '" size="25" value="" onchange="valThis(this,\'' + rsVars[i].varVar + '\');">' +
								'<select size="1" name="valValueF' + rsVars[i].varVar + '" id="valValueF' + rsVars[i].varVar + '">' +
								'<option></option>' +
								'<option value="@LogNum">' + DtxtLogNum + '</option>' +
								'<option value="@LogNum">' + DtxtLineNum + '</option>' +
								'<option value="@ItemCode">' + DtxtItemCode + '</option>' +
								'<option value="@PriceList">' + DtxtPList + '</option>' +
								'<option value="@CardCode">' + DtxtClientCode + '</option>' +
								'<option value="@SlpCode">' + DtxtAgentCode + '</option>' +
								'<option value="@dbName">' + DtxtDB + '</option>' +
								'<option value="@WhsCode">' + DtxtWhsCode + '</option>' +
								'<option value="@Quantity">' + DtxtQty + '</option>' +
								'<option value="@Unit">' + DtxtUnit + '</option>' +
								'<option value="@Price">' + DtxtPrice + '</option>' +
								'</select></font>' +
								'<table cellpadding="0" cellspacing="0" border="0" width="100%" style="display: none;" id="tblQuery' + rsVars[i].varVar + '">' +
								'	<tr>' +
								'		<td>' +
								'			<textarea dir="ltr" style="width: 100%; font-size: 10px; font-family: Verdana; color: #3F7B96; font-weight: bold; border: 1px solid #68A6C0; background-color: #D9F0FD" name="valQuery' + rsVars[i].varVar + '" id="valQuery' + rsVars[i].varVar + '" onchange="javascript:document.form2.btnVerfyQueryVar' + rsVars[i].varVar + '.src=\'images/btnValidate.gif\';document.form2.btnVerfyQueryVar' + rsVars[i].varVar + '.style.cursor = \'hand\';document.form2.valQueryVar' + rsVars[i].varVar + '.value=\'Y\';"></textarea>' +
								'		</td>' +
								'		<td width="1" valign="bottom">' +
								'			<img src="images/btnValidateDis.gif" id="btnVerfyQueryVar' + rsVars[i].varVar + '" title="' + DtxtValidate + '" onclick="javascript:if (document.form2.valQueryVar' + rsVars[i].varVar + '.value == \'Y\')VerfyQueryVar(\'' + rsVars[i].varVar + '\');">' +
								'			<input type="hidden" name="valQueryVar' + rsVars[i].varVar + '" value="N">' +
								'		<img src="images/spacer.gif"></td>'+
								'	</tr>' +
								'</table>' + 
								'<input type="hidden" name="varDataType' + rsVars[i].varVar + '" id="varDataType' + rsVars[i].varVar + '" value="' + rsVars[i].varDataType + '">' +
								'<input type="hidden" name="varVar" value="' + rsVars[i].varVar + '">' +
								'<input type="hidden" name="varNotNull" value="' + rsVars[i].varNotNull + '">';
			if (rsVars[i].varDataType == 'datetime')
			{
				Calendar.setup({
					inputField     :    "colValDat" + rsVars[i].varVar,     // id of the input field
					ifFormat       :    CalendarFormat,      // format of the input field
					button         :    "btnValDatImg" + rsVars[i].varVar,  // trigger for the calendar (button ID)
					align          :    "Bl",           // alignment (defaults to "Bl")
					singleClick    :    true
				});
			}
		}
	}
}
function valFrm2()
{
	if (document.form2.valQuery.value == 'Y')
	{
		alert(LtxtValQryVal);
		document.form2.btnVerfyFilter.focus();
		return false;
	}
	else if (document.form2.Name.value == '')
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
	else if (document.form2.linkActive.checked && document.form2.linkObject.selectedIndex < 1)
	{
		alert(LtxtValRepLnk);
		document.form2.linkObject.focus();
		return false;
	}
	else if (rsHasVars)
	{
		if (document.form2.varVar.length)
		{
			for (var i = 0;i<document.form2.varVar.length;i++)
			{
				if (!valFrmVar(document.form2.varVar[i].value, document.form2.varNotNull[i].value)) return false;
			}
		}
		else
		{
			if (!valFrmVar(document.form2.varVar.value, document.form2.varNotNull.value)) return false;
		}
	}
	return true;
}

function changeValBy(varVar, by)
{
	switch (by)
	{
		case 'F':
			document.getElementById('valValueV' + varVar).style.display='none';
			document.getElementById('tblValDat' + varVar).style.display='none';
			document.getElementById('tblQuery' + varVar).style.display='none';
			document.getElementById('valValueF' + varVar).style.display='';
			break;
		case 'V':
			if (document.getElementById('varDataType' + varVar).value != 'datetime')
			{
				document.getElementById('valValueV' + varVar).style.display='';
			}
			else
			{
				document.getElementById('tblValDat' + varVar).style.display='';
			}
			document.getElementById('valValueF' + varVar).style.display='none';
			document.getElementById('tblQuery' + varVar).style.display='none';
			break;
		case 'Q':
			document.getElementById('tblValDat' + varVar).style.display='none';
			document.getElementById('valValueV' + varVar).style.display='none';
			document.getElementById('tblQuery' + varVar).style.display='';
			document.getElementById('valValueF' + varVar).style.display='none';
			break;
	}
}

function valFrmVar(varVar, varNotNull)
{
	if (varNotNull == 'Y')
	{
		var rdFld = document.getElementById('rdFld' + varVar);
		var rdVal = document.getElementById('rdVal' + varVar);
		var rdQry = document.getElementById('rdQry' + varVar);
		
		if (rdFld.checked)
		{
			if (document.getElementById('valValueF' + varVar).selectedIndex == 0)
			{
				alert(LtxtSelFldVar.replace('{0}', varVar));
				return false;
			}
		}
		else if (rdVal.checked)
		{
			if (document.getElementById('varDataType' + varVar).value != 'datetime')
			{
				if (document.getElementById('valValueV' + varVar).value == '')
				{
					alert(LtxtValVarVal.replace('{0}', varVar));
					return false;
				}
			}
			else
			{
				if (document.getElementById('colValDat' + varVar).value == '')
				{
					alert(LtxtValVarVal.replace('{0}', varVar));
					return false;
				}
			}
		}
		else
		{
			if (document.getElementById('valQuery' + varVar).value == '')
			{
				alert(LtxtValVarQry.replace('{0}', varVar));
				document.getElementById('valQuery' + varVar).focus();
				return false;
			}
			else if (document.getElementById('valQueryVar' + varVar).value == 'Y')
			{
				alert(LtxtValVarQryBrnVerfy.replace('{0}', varVar));
				document.getElementById('btnVerfyQueryVar' + varVar).focus();
				return false;
			}
		}
		return true;
	}
	else return true;
}

function valThis(fld, varVar)
{
	varDataType = document.getElementById('varDataType' + varVar).value;
	if ((varDataType == 'float' || varDataType == 'numeric' || varDataType == 'int') && fld.value != '')
	{
		if (!IsNumeric(fld.value))
		{
			alert(DtxtValNumVal);
			fld.value = '';
			fld.focus();
		}
	}
}

var myBtnVerfy;
var myHdVerfy;
function VerfyQuery()
{
	document.frmVerfyQuery.Query.value = document.form2.customSql.value;
	myBtnVerfy = document.form2.btnVerfyFilter;
	myHdVerfy = document.form2.valQuery;
	document.frmVerfyQuery.submit();
}

function VerfyQueryVerified()
{
	//myBtnVerfy.disabled = true;
	myBtnVerfy.src='images/btnValidateDis.gif'
	myBtnVerfy.style.cursor = '';
	myHdVerfy.value='N';
}
function VerfyQueryVar(varVar)
{
	myBtnVerfy = document.getElementById('btnVerfyQueryVar' + varVar);
	myHdVerfy = document.getElementById('valQueryVar' + varVar);
	document.frmVerfyQuery.Query.value = document.getElementById('valQuery' + varVar).value;
	document.frmVerfyQuery.submit();
}