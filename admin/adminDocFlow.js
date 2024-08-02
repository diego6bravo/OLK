function showTab(tabID)
{
	var myTabs = new Array('tabQry', 'tabMsg', 'tabLines', 'tabAutGrp');
	for (var i = 0;i<myTabs.length;i++)
	{
		if (myTabs[i] == tabID)
		{
			document.getElementById('btn' + tabID).style.backgroundColor = '#D9F5FF';
			document.getElementById('btn' + tabID).style.cursor = '';
			document.getElementById(tabID).style.display = '';
		}
		else
		{
			document.getElementById('btn' + myTabs[i]).style.backgroundColor = '#BFEEFE';
			document.getElementById('btn' + myTabs[i]).style.cursor = 'hand';
			document.getElementById(myTabs[i]).style.display = 'none';
		}
	}
}

function changeApplyToClient()
{
/*	if (document.frmAddEditFlow.FlowType.value == '2' && document.frmAddEditFlow.ApplyToClient.checked)
	{
		alert(txtChangeApplyToClie);
	}
	*/
}
function changeFlow(val)
{
	ExecAt = document.frmAddEditFlow.ExecAt.value;
	if (val == '2' && ExecAt != 'D3' && ExecAt != 'R2' && ExecAt != 'A1' && ExecAt != 'C1' && Left(ExecAt, 1) != 'O' && ExecAt != '')
	{
		alert(txtFlowComExec.replace('{0}', document.frmAddEditFlow.ExecAt.options[document.frmAddEditFlow.ExecAt.selectedIndex].text));
		document.frmAddEditFlow.FlowType.selectedIndex = 1;
	}
	/*else if (val == '2' && document.frmAddEditFlow.ApplyToClient.checked && Left(ExecAt, 1) != 'O')
	{
		alert(txtFlowTypeCAlr);
	}*/
	else if (val == '1' && document.frmAddEditFlow.ApplyToClient.checked && ExecAt == 'C1')
	{
		alert(txtFlowTypeCAlr2);
	}
	activateControls();
}
function changeDoc(val)
{
	if (val != 'D3' && val != 'R2' && val != 'A1' && val != 'C1' && Left(val,1) != 'O' && val != '' && document.frmAddEditFlow.FlowType.value == '2')
	{
		alert(txtExecComFlow.replace('{0}', document.frmAddEditFlow.ExecAt.options[document.frmAddEditFlow.ExecAt.selectedIndex].text));
		document.frmAddEditFlow.ExecAt.selectedIndex = 0;
	}
	
	enObj = false
	if (Left(val,1) != 'D') enObj = true;
	for (var i = 1;i<=10;i++) { document.getElementById('ObjectCode' + i).disabled = enObj; }
	
	document.frmAddEditFlow.ApplyToClient.disabled = !(val == 'D2' || val == 'D3' || val == 'C1' || Left(val, 1) == 'O');
	
	if (val != '')
	{
		myVars = '';
		if (val != 'D1' && val != 'R1' && Left(val, 1) != 'O') myVars = '<span dir="ltr">@LogNum</span> = ' + txtOLKDocKey + '<br>';
		if (val == 'O2' || val == 'O3' || val == 'O4') myVars += '<span dir="ltr">@ObjectCode</span> = ' + txtObjCode + '<br>';
		if (Left(val, 1) == 'O' && Left(val, 2) != 'OP') myVars += '<span dir="ltr">@Entry</span> = ' + txtDocEntry + '<br>';
		if (Left(val, 2) == 'OP')
		{
			myVars += '<span dir="ltr">@LogNum</span> = ' + txtOLKDocKey + '<br>';
		}
		
		myVars += 	'<span dir="ltr">@LanID</span> = ' + txtLanID + '<br>' +
					'<span dir="ltr">@SlpCode</span> = ' + txtAgentCode + '<br>' +
					'<span dir="ltr">@dbName</span> = ' + txtDB + '<br>' +
					'<span dir="ltr">@branch</span> = ' + txtBranchCode;
					
		if (Left(val,1) == 'D' || Left(val,1) == 'R' || val == 'C2' || val == 'C3') myVars += '<br><span dir="ltr">@CardCode</span> = ' + txtClientCode;
		
		if (val == 'D2') 
		{
			myVars += '<br><span dir="ltr">@ItemCode</span> = ' + txtItemCode;
			myVars += '<br><span dir="ltr">@WhsCode</span> = ' + txtWhsCode;
			myVars += '<br><span dir="ltr">@Quantity</span> = ' + txtQtyInUnit;
			myVars += '<br><span dir="ltr">@Unit</span> = ' + txtUnit;
			myVars += '<br><span dir="ltr">@Price</span> = ' + txtPrice;

		}
		document.getElementById('txtVars').innerHTML = myVars;
	}
	else
	{
		document.getElementById('txtVars').innerHTML = "";
	}
	
	if (document.frmAddEditFlow.FlowQuery.value != '') document.frmAddEditFlow.btnVerfyFlowQuery.disabled = false;
	if (document.frmAddEditFlow.LineQuery.value != '') document.frmAddEditFlow.btnVerfyLineQuery.disabled = false;
	if (document.frmAddEditFlow.NoteQuery.value != '') document.frmAddEditFlow.btnVerfyNoteQuery.disabled = false;
	
	activateControls();
}
function activateControls()
{
	if (document.frmAddEditFlow.ApplyToClient.disabled)document.frmAddEditFlow.ApplyToClient.checked=false;
	
	document.frmAddEditFlow.FlowDraft.disabled = document.frmAddEditFlow.FlowType.value != 1 || document.frmAddEditFlow.FlowType.value == 1 && document.frmAddEditFlow.ExecAt.value != 'D3' && document.frmAddEditFlow.ExecAt.value != 'R2';
	document.frmAddEditFlow.FlowAuthorize.disabled = document.frmAddEditFlow.FlowType.value != 1 || document.frmAddEditFlow.FlowType.value == 1 && document.frmAddEditFlow.ExecAt.value != 'D3';
}

function valFrm2()
{
	if (document.frmAddEditFlow.FlowName.value == '')
	{
		alert(txtValFlowNam);
		document.frmAddEditFlow.FlowName.focus();
		return false;
	}
	else if (document.frmAddEditFlow.SlpCode.value == '' && document.frmAddEditFlow.ApplyToClient.disabled)
	{
		alert(txtValSelAgent);
		document.frmAddEditFlow.Agents.focus();
		return false;
	}
	else if (document.frmAddEditFlow.SlpCode.value == '' && (!document.frmAddEditFlow.ApplyToClient.disabled && !document.frmAddEditFlow.ApplyToClient.checked))
	{
		alert(txtValSelAgentOrClie);
		document.frmAddEditFlow.Agents.focus();
		return false;
	}
	else if (document.frmAddEditFlow.ExecAt.selectedIndex == 0)
	{
		alert(txtValExecMom);
		document.frmAddEditFlow.ExecAt.focus();
		return false;
	}
	else if (Left(document.frmAddEditFlow.ExecAt.value,1) == 'D' && !chkDocSelected())
	{
		alert(txtValSelDoc);
		return false;
	}
	else if (document.frmAddEditFlow.FlowQuery.value == '')
	{
		alert(txtValFlowQry);
		showTab('tabQry');
		document.frmAddEditFlow.FlowQuery.focus();
		return false;
	}
	else if (document.frmAddEditFlow.valFlowQuery.value == 'Y')
	{
		alert(txtVarFlowQryVal);
		showTab('tabQry');
		document.frmAddEditFlow.btnVerfyFlowQuery.focus();
		return false;
	}
		else if (document.frmAddEditFlow.NoteQuery.value == '' && document.frmAddEditFlow.NoteBuilder.checked)
	{
		alert(txtValMsgQry);
		showTab('tabMsg');
		document.frmAddEditFlow.NoteQuery.focus();
		return false;
	}
	else if (document.frmAddEditFlow.NoteQuery.value != '' && document.frmAddEditFlow.valNoteQuery.value == 'Y')
	{
		alert(txtValMsgQryVal);
		showTab('tabMsg');
		document.frmAddEditFlow.btnVerfyNoteQuery.focus();
		return false;
	}
	else if (document.frmAddEditFlow.NoteText.value == '')
	{
		alert(txtValMsgText);
		showTab('tabMsg');
		document.frmAddEditFlow.NoteText.focus();
		return false;
	}
	else if (document.frmAddEditFlow.LineQuery.value != '' && document.frmAddEditFlow.valLineQuery.value == 'Y')
	{
		alert(txtValLineQryVal);
		showTab('tabLines');
		document.frmAddEditFlow.btnVerfyLineQuery.focus();
		return false;
	}
	
	if (document.frmAddEditFlow.GrpID)
	{
		if (document.frmAddEditFlow.GrpID.length)
		{
			for (var i = 0;i<document.frmAddEditFlow.GrpID.length;i++)
			{
				var grpID = document.frmAddEditFlow.GrpID[i].value;
				
				if (document.getElementById('valGrpValue' + grpID + 'Query').value == 'Y' && document.getElementById('GrpQuery' + grpID).checked)
				{
					alert(txtVarAutGrpQry);
					showTab('tabAutGrp');
					document.getElementById('GrpValue' + grpID + 'Query').focus();
					return false;
				}
			}
		}
		else
		{
			var grpID = document.frmAddEditFlow.GrpID.value;
			
			if (document.getElementById('valGrpValue' + grpID + 'Query').value == 'Y' && document.getElementById('GrpQuery' + grpID).checked)
			{
				alert(txtVarAutGrpQry);
				showTab('tabAutGrp');
				document.getElementById('GrpValue' + grpID + 'Query').focus();
				return false;
			}
		}
	}
	return true;
}
function chkDocSelected()
{
	if (!document.frmAddEditFlow.FlowActive.checked) return true;
	
	for (var i = 1;i<=10;i++)
	{
		if (document.getElementById('ObjectCode' + i).checked) { return true; }
	}
	return false;
}
var setBy
function VerfyQuery(By)
{
	if (document.frmAddEditFlow.ExecAt.selectedIndex != 0)
	{
		setBy = By;
		document.frmVerfyQuery.Query.value = document.getElementById(By + 'Query').value;
		document.frmVerfyQuery.by.value = By;
		document.frmVerfyQuery.ExecAt.value = document.frmAddEditFlow.ExecAt.value;
		document.frmVerfyQuery.submit();
	}
	else
	{
		alert(txtValExecMom);
		document.frmAddEditFlow.ExecAt.focus();
	}
}
function VerfyQueryVerified()
{
	document.getElementById('btnVerfy' + setBy + 'Query').src='images/btnValidateDis.gif';
	document.getElementById('btnVerfy' + setBy + 'Query').style.cursor = '';
	document.getElementById('val' + setBy + 'Query').value='N';
}

function ChkNoteBuilder()
{
	if(document.frmAddEditFlow.NoteBuilder.checked)
	{
		if (document.frmAddEditFlow.NoteQuery.value != '')
		{
			document.frmAddEditFlow.btnVerfyNoteQuery.disabled=false;
		}
		document.frmAddEditFlow.NoteQuery.disabled=false;
	}
	else
	{
		document.frmAddEditFlow.btnVerfyNoteQuery.disabled=true;
		document.frmAddEditFlow.NoteQuery.disabled=true;
	}
	document.frmAddEditFlow.NoteQueryFields.disabled = document.frmAddEditFlow.NoteQuery.disabled;
}
function getQueryFields() { return document.frmAddEditFlow.NoteQueryFields; }
function selectAgents()
{
	OpenWin = window.open('selectAgents.asp?SlpCode='+document.frmAddEditFlow.SlpCode.value+'&pop=Y','OpenWin', 'width=320,height=480,scrollbars=yes');
}
function agentsToCode(SlpCode, SlpName) 
{ 
	document.frmAddEditFlow.SlpCode.value = SlpCode; 
	document.frmAddEditFlow.Agents.value = SlpName;
}


var maxOrdr = 0;

function doAddGrp(grp)
{
	var grpID = grp.value;
	var grpName = grp.options[grp.selectedIndex].innerText;
	
	var str = '';
	str += '<tr>' ;
	str += '<td class="TblRepNrm">' ;
	str += '<input type="hidden" name="GrpID" value="' + grpID + '">' + grpName + '</td>' ;
	str += '<td class="TblRepNrm"><input type="checkbox" name="AsignedSLP' + grpID + '" id="AsignedSLP' + grpID + '" value="Y" class="noborder"><label for="AsignedSLP' + grpID + '">' + txtAsignedSLP + '</label></td>' ;
	str += '<td class="TblRepNrm"><input type="checkbox" name="GrpQuery' + grpID + '" id="GrpQuery' + grpID + '" value="Y" class="noborder" onclick="document.getElementById(\'trQryGrpID' + grpID + '\').style.display=this.checked?\'\':\'none\';"><label for="GrpQuery' + grpID + '">' + txtAlertFilter + '</label></td>' ;
	str += '<td align="center"><table cellpadding="0" border="0">' ;
	str += '<tr>' ;
	str += '<td class="TblRepNrm"><input type="text" name="Order' + grpID + '" id="Order' + grpID + '" size="5" style="text-align:right" class="input" value="' + (maxOrdr++) + '" onfocus="this.select()" onkeydown="return chkMax(event, this, 6);"></td>' ;
	str += '<td valign="middle">' ;
	str += '<table cellpadding="0" cellspacing="0" border="0">' ;
	str += '<tr>' ;
	str += '<td><img src="images/img_nud_up.gif" id="btnOrder' + grpID + 'Up"></td>' ;
	str += '</tr>' ;
	str += '<tr>' ;
	str += '<td><img src="images/spacer.gif"></td>' ;
	str += '</tr>' ;
	str += '<tr>' ;
	str += '<td><img src="images/img_nud_down.gif" id="btnOrder' + grpID + 'Down"></td>' ;
	str += '</tr>' ;
	str += '</table>' ;
	str += '</td>' ;
	str += '</tr>' ;
	str += '</table>' ;
	str += '<script language="javascript">NumUDAttach(\'frmAddEditFlow\', \'Order' + grpID + '\', \'btnOrder' + grpID + 'Up\', \'btnOrder' + grpID + 'Down\');</script>' ;
	str += '</td>' ;
	str += '<td class="TblRepNrm"><img border="0" src="images/remove.gif" width="16" height="16" onclick="delGrp(this, ' + grpID + ');"></td>' ;
	str += '</tr>' ;
	str += '<tr id="trQryGrpID' + grpID + '" style="display: none;">' ;
	str += '<td colspan="5">' ;
	str += '<font face="Verdana" size="1" color="#4783C5">where SlpCode in (...)</font>' ;
	str += '<table cellpadding="0" cellspacing="0" border="0" width="100%">' ;
	str += '<tr>' ;
	str += '<td rowspan="2">' ;
	str += '<textarea dir="ltr" rows="10" style="width: 100%" name="GrpValue' + grpID + 'Query" cols="100" class="input" onkeypress="javascript:document.frmAddEditFlow.btnVerfyGrpValue' + grpID + 'Query.src=\'images/btnValidate.gif\';document.frmAddEditFlow.btnVerfyGrpValue' + grpID + 'Query.style.cursor = \'hand\';document.frmAddEditFlow.valGrpValue' + grpID + 'Query.value=\'Y\';"></textarea>' ;
	str += '</td>' ;
	str += '<td valign="top" width="1">' ;
	str += '<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteGrpFilter" alt="|D:txtDefinition|" onclick="javascript:doFldNote(11, \'GrpValue' + grpID + 'Query\', \'\', null);">' ;
	str += '</td>' ;
	str += '</tr>' ;
	str += '<tr>' ;
	str += '<td valign="bottom" width="1">' ;
	str += '<img src="images/btnValidateDis.gif" id="btnVerfyGrpValue' + grpID + 'Query" alt="|D:txtValidate|" onclick="javascript:if (document.frmAddEditFlow.valGrpValue' + grpID + 'Query.value == \'Y\')VerfyQuery(\'GrpValue' + grpID + '\');">' ;
	str += '<input type="hidden" name="valGrpValue' + grpID + 'Query" value="N">' ;
	str += '</td>' ;
	str += '</tr>' ;
	str += '</table>' ;
	str += '</td>' ;
	str += '</tr>' ;

	$("#tblAutGrp tr:last").before(str);

	NumUDAttach('frmAddEditFlow', 'Order' + grpID, 'btnOrder' + grpID + 'Up', 'btnOrder' + grpID + 'Down');
	
	grp.options.remove(grp.selectedIndex);
	grp.selectedIndex = 0;
	
}

function delGrp(grp, delGrpCode)
{
	if (confirm(txtConfDel))
	{
		$(grp).parent().parent().remove();
		
		$('#trQryGrpID' + delGrpCode).remove();
		
		var delID = document.frmAddEditFlow.delID;
		if (delID.value != '') delID.value += ', ';
		delID.value += delGrpCode;

	}
}