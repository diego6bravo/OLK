  jQuery(document).ready(function() {
    jQuery("#flowViewDialog").dialog({
      bgiframe: true, autoOpen: false, width: 760, height: 400, minWidth: 550, minHeight: 400, modal: true, resizable: false
    });
  });


function viewFlowLog(id)
{
	flowObjDesc = document.getElementById('ObjDesc' + id).value;
	
	$.post("flowViewControlFetch.asp?d=" + (new Date()).toString(), { ID: id, Type: 'L' },
   function(data){
     doViewFlowResult(data);
   });

}

var firstFlowViewID;
var flowObjDesc;
function doViewFlowResult(data)
{
	if (data)
	{
		try
		{
			arrData = data.split('{L}');
			var strViewFlow = 	'<table style="width: 100%" cellpadding="0" cellspacing="1" border="0"> ' +  
									'<tr class="GeneralTlt" style="text-align: center;"> ' +  
										'<td colspan="2">' + txtReqDate + '</td> ' +  
										'<td>' + txtReqUser + '</td> ' +  
										'<td>' + txtRespDate + '</td> ' +  
										'<td>' + txtRespUser + '</td> ' +  
										'<td>&nbsp;</td> ' +  
									'</tr> '; 	

			for (var i = 0;i<arrData.length;i++)
			{
				arrValues = arrData[i].split('{S}');
				var id = arrValues[0];
				if (i == 0) firstFlowViewID = id;
				var requestDate = arrValues[1];
				var requestUser = arrValues[2];
				var confirmDate = arrValues[3];
				var confirmUser = arrValues[4];
				var confirmNote = arrValues[5];
				var status = arrValues[6];
				
				var statusStr;
				
				switch (status)
				{
					case 'C':
						statusStr = txtCanceled;
						break;
					case 'R':
						statusStr = txtRejected;
						break;
					case 'X':
						statusStr = txtReOpened;
						break;
					case 'S':
						statusStr = txtProcessed;
						break;
					case 'O':
						statusStr = txtWaiting;
						break;

				}
				
				strViewFlow += 	'<tr class="GeneralTblBold2" style="cursor: pointer; " onclick="showViewDetail(' + id + ');"> ' +  
									'<td width="8"><img src="images/' + rtl + 'arrows.gif" id="flowImg' + id + '"></td>' +
									'<td align="center" dir="ltr">' + requestDate + '</td> ' + 
									'<td>' + requestUser + '</td> ' +  
									'<td align="center" dir="ltr">' + confirmDate + '</td> ' +  
									'<td>' + confirmUser + '</td> ' +  
									'<td align="center">' + statusStr + '</td> ' +  
								'</tr> ' +
								'<tr id="flowViewDetailRow' + id + '" style="display: none; ">' +  
									'<td colspan="6" id="flowViewDetail' + id + '">' +  
									'</td>' +  
								'</tr>';
								
				if (confirmNote != '')
				{
					strViewFlow += 	'<tr id="flowViewNote' + id + '" style="display: none; ">' +  
										'<td colspan="6">' +  
											'<table style="width: 100%"><tr>' +
											'<td class="GeneralTblBold2HighLight" width="120">' + txtRespNote + '</td>' +
											'<td class="FirmTblY">' + confirmNote + '</td>'+
											'</tr></table>';
										'</td>' +  
									'</tr>' 
				}
			}
			
			strViewFlow += '</table> ';
			
			document.getElementById('flowViewDialogContent').innerHTML = strViewFlow;
			
			showViewDetail(firstFlowViewID);
			
			jQuery('#flowViewDialog').dialog('option', 'title', txtFlowViewControl + ' - #' + firstFlowViewID + ' - ' + flowObjDesc);
			jQuery('#flowViewDialog').dialog('open');
			
			setTimeout('$(\'#flowViewDialogContent\').scrollTop(0);', 0);


		}
		catch(err)
		{
		}
	}
}

var flowViewDetail;
function showViewDetail(id)
{
	flowViewDetail = document.getElementById('flowViewDetail' + id);
	var flowViewNote = document.getElementById('flowViewNote' + id);
	var flowViewDetailRow = document.getElementById('flowViewDetailRow' + id);
	var flowImg = document.getElementById('flowImg' + id);
	
	if (flowViewDetail.innerHTML == '')
	{
		
		$.post("flowViewControlFetch.asp?d=" + (new Date()).toString(), { ID: id, Type: 'D' },
	   function(data){
	     doShowViewDetail(data);
	   });
	}
	
	flowViewDetailRow.style.display = flowViewDetailRow.style.display == 'none' ? '' : 'none';
	if (flowViewNote) flowViewNote.style.display = flowViewDetailRow.style.display;
	flowImg.src = 'images/' + (flowViewDetailRow.style.display == 'none' ? rtl : '') + (flowViewDetailRow.style.display == 'none' ? 'arrows' : 'arrow_down') + '.gif';
}

function doShowViewDetail(data)
{
	arrFlow = data.split('{F}');
	
	var strDetails = '';
	
	for (var i = 0;i<arrFlow.length;i++)
	{
		var arrData = arrFlow[i].split('{S}');
		var flowName = arrData[0];
		var flowTable = arrData[1];
		var flowNote = arrData[2];
		var userNote = arrData[3];
		strDetails += '<table style="width: 100%">';
		strDetails += '<tr>';
		strDetails += '<td class="MsgTlt2" width="120">' + txtFlow + '</td>';
		strDetails += '<td class="GeneralTbl"><b><i>' + flowName + '<b><i></td>';
		strDetails += '</tr>';
		strDetails += '<tr>';
		strDetails += '<td class="MsgTlt2" width="120">' + txtFlowNote + '</td>';
		strDetails += '<td class="FirmTblY">' + flowNote + '</td>';
		strDetails += '</tr>';
		
		if (userNote != '')
		{
			strDetails += '<tr>';
			strDetails += '<td class="MsgTlt2" width="120">' + txtUserReq + '</td>';
			strDetails += '<td class="FirmTblY">' + userNote + '</td>';
			strDetails += '</tr>';
		}
		
		strDetails += '</table>';

	}
	flowViewDetail.innerHTML = strDetails;
	
}