var flowFunc = '';
var flowExec = '';
var loadVar = '';
var flowRedir = '';
var typeIDs = '';
var typeFlowIDs = '';
var flowDraftFld;
var flowAutFld;

  jQuery(document).ready(function() {
    jQuery("#flowDialog").dialog({
      bgiframe: true, autoOpen: false, width: 700, height: 300, minWidth: 700, minHeight: 300, modal: true, resizable: false
    });
  });
  
function processFlow()
{
	var flowNotes = '';
	var arrFlow = typeFlowIDs.split(', ');
	for (var i = 0;i<arrFlow.length;i++)
	{
		if (i > 0) flowNotes += '{S}';
		flowNotes += document.getElementById('FlowNote' + arrFlow[i]).value;
	}
	$.post("flowControlProcess.asp?d=" + (new Date()).toString(), { ExecAt: flowExec, Variables: loadVar, FlowID: typeFlowIDs, FlowNotes: flowNotes },
   function(data)
   {
		if (data == 'ok')
		{
			alert(txtConfFlow);
			closeFlowAlert();
			if (flowRedir != '') window.location.href = flowRedir;
		}
		else
		{
			alert('An error ocurred while processing flow: ' + data);
		}
   });
}
function setFlowAlertVars(execAt, variables, flowFunction, redirectValue)
{
	flowExec = execAt;
	loadVar = variables;
	flowFunc = flowFunction;
	flowRedir = redirectValue;
}

function doFlowAlert()
{
	typeIDs = '';
	typeFlowIDs = '';
	
	$.post("flowControlFetch.asp?d=" + (new Date()).toString(), { ExecAt: flowExec, Variables: loadVar },
   function(data){
     doFlowAlertResult(data);
   });
}
function doFlowAlertResult(data)
{
	if (data != '' && data != '1')
	{
		try
		{
			arrFlow = data.split('{F}');
			if (arrFlow.length > 0)
			{
				var strFlow = '';
				var imgAlign = rtl == '' ? 'right' : 'left';
				
				var flowType;
				for (var i = 0;i<arrFlow.length;i++)
				{
					arrValues = arrFlow[i].split('{S}');
					var flowID = arrValues[0];
					var flowName = arrValues[1];
					flowType = parseInt(arrValues[2]);
					var flowTable = arrValues[3];
					var note = arrValues[4];
					var draft = arrValues[5];
					var authorize = arrValues[6];
					
					var strImg = '';
					var strImgAlt = '';
					var btnOk = '';
					var btnCancel = '';
					
					if (flowType == 2)
					{
						if (typeFlowIDs != '') typeFlowIDs += ', ';
						typeFlowIDs = typeFlowIDs + flowID;
					}
					
					if (typeIDs != '') typeIDs += ', ';
					typeIDs = typeIDs + flowID;
				
					switch (flowType)
					{
						case 0:
							strImg = 'error';
							strImgAlt = altError;
							btnCancel = txtClose;
							break;
						case 1:
							strImg = 'confirm';
							strImgAlt = altConf;
							btnOk = txtContinue;
							btnCancel = txtCancel;
							break;
						case 2:
							strImg = 'question';
							strImgAlt = altFlow;
							btnOk = txtConfirm;
							btnCancel = txtCancel;
							break;
					}
					
					strFlow +=	 		'<div class="GeneralTlt">' + flowName + '</div>' +  
										'<div class="GeneralTbl" style="margin: 1px; height: 100px; overflow: auto;"> ' +  
										'<img border="0" src="design/' + SelDes + '/images/' + strImg + 'icon.gif" width="68" height="65" alt="' + strImgAlt + '" style="float: ' + imgAlign + '">' + note + ' ' +  
										'</div>';
										
					if (flowType == 2)
						strFlow += 		'<div>' +  
										'<table cellpadding="0" cellspacing="0" border="0" style="width: 100%;">' +  
										'	<tr class="GeneralTblBold">' +  
										'		<td width="60px"><nobr>' + txtNote + ':</nobr></td>' +  
										'		<td><input class="input" type="text" id="FlowNote' + flowID + '" name="FlowNote' + flowID + '" maxlength="256" onkeydown="doExecFlowOK(event);" style="width: 100%;"></td>' +  
										'	</tr>' +  
										'</table>' +  
										'</div>'
	
					if (flowTable != '')
					{
						var arrRows = flowTable.split('{R}');
						
						strFlow += 		'<div style="height: 100px; overflow: auto; margin-bottom: 1px;">' +  
										'<table style="width: 100%">';
										
						for (var r = 0;r<arrRows.length;r++)
						{
							var arrCols = arrRows[r].split('{C}');
							var rowCSS = r > 0 ? 'GeneralTbl' : 'GeneralTlt';
							
							strFlow += '	<tr class="' + rowCSS + '">';
							
							for (var c = 0;c<arrCols.length;c++)
							{
								strFlow += '<td>' + arrCols[c] + '</td>';
							}
							
							strFlow += '	</tr>';
						}
						
						strFlow += 		'</table>' +  
										'</div>';
					}
					
					strFlow += '<br>';
				}
				strFlow += 				'<div style="text-align: center; margin-bottom: 1px; margin-top: 1px;">';
				
				if (btnOk != '') 
				{
					var finalFunc = flowFunc;
					switch(flowType)
					{
						case 1:
							if (draft == 'Y')
							{
								if (flowDraftFld != null && flowDraftFld != '') finalFunc = flowDraftFld + '=\'Y\';' + finalFunc;
							}
							if (authorize == 'Y')
							{
								if (flowAutFld != null && flowAutFld != '') finalFunc = flowAutFld + '=\'Y\';' + finalFunc;
							}
							break;
						case 2:
							finalFunc = 'processFlow();';
							break;
					}
	
					strFlow += 			'<input type="button" value="' + btnOk + '" name="btnFlowOk" id="btnFlowOk" onclick="javascript:this.disabled=true;' + finalFunc + '" class="ui-button ui-state-default ui-corner-all" style="height: 30px; width: 120px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;';
				}
					
				strFlow += 				'<input type="button" value="' + btnCancel + '" name="btnCancel" onclick="javascript:closeFlowAlert();" class="ui-button ui-state-default ui-corner-all" style="height: 30px; width: 120px;">' +  
										'</div>' ;
										
				document.getElementById('flowDialogContent').innerHTML = strFlow;
								
				jQuery('#flowDialog').dialog('open');
				
				setTimeout('$(\'#flowDialogContent\').scrollTop(0);', 0);
			}
			else
			{
			}
		}	
		catch (err)
		{
			alert('Flow Alert Result Error: ' + err.description);
		}
	}
	else
	{
		setTimeout(flowFunc, 0);
	}
	
}
function closeFlowAlert()
{
	jQuery('#flowDialog').dialog('close');
}

function doExecFlowOK(e)
{
	if (e.keyCode == 13) document.getElementById('btnFlowOk').click();
}
