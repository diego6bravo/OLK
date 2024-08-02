function Start(page) 
{
   var wOpen;
   var sOptions;

   sOptions = 'toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes';
   sOptions = sOptions + ',width=' + (screen.availWidth - 10).toString();
   sOptions = sOptions + ',height=' + (screen.availHeight - 122).toString();
   sOptions = sOptions + ',screenX=0,screenY=0,left=0,top=0';

   wOpen = window.open('', 'objDetails', sOptions );
   wOpen.location = page;
   wOpen.focus();
   wOpen.moveTo( 0, 0 );
   wOpen.resizeTo( screen.availWidth, screen.availHeight );
   OpenWin = wOpen;
   
}
function goBP(code)
{
	Start('');
	doMyLink('addCard/crdConfDetailOpen.asp', 'CardCode=' + code + '&pop=Y&AddPath=', 'objDetails');
}
function goCXC(cardCode)
{
	Start('');
	doMyLink('cxcPrint.asp', 'c1=' + cardCode + '&SumDec=' + SumDec + '&LinkRep=Y&pop=Y&AddPath=', 'objDetails');
}
function goDetail(execAt, objCode, objEntry)
{
	if (Left(execAt, 1) == 'O')
	{
		switch (objCode)
		{
			case 2:
				Start('');
				doMyLink('addCard/crdConfDetailOpen.asp', 'CardCode=' + objEntry + '&pop=Y&AddPath=', 'objDetails');
				return;
			case 4:
				Start('');
				doMyLink('addItem/itmConfDetail.asp', 'ItemCode=' + objEntry + '&pop=Y&AddPath=', 'objDetails');
				return;
			case 24:
			case 46:
				Start('');
				doMyLink('cxcRctDetailOpen.asp', 'isEntry=Y&DocType=' + objCode + '&DocEntry=' + objEntry + '&pop=Y', 'objDetails');
				return;
			case 33:
				Start('');
				doMyLink('addActivity/activityConfDetail.asp', 'isEntry=Y&DocType=' + objCode + '&DocEntry=' + objEntry + '&pop=Y&AddPath=', 'objDetails');
				return;
			case 13:
			case 14:
			case 15:
			case 16:
			case 17:
			case 18:
			case 19:
			case 20:
			case 21:
			case 22:
			case 23:
				Start('');
				doMyLink('cxcDocDetailOpen.asp', 'isEntry=Y&DocType=' + objCode + '&DocEntry=' + objEntry + '&pop=Y', 'objDetails');
				return;
			default:
				alert('Object not supported');
				return;
		}
	}
	
	switch (execAt)
	{
		case 'C1':
			Start('');
			doMyLink('addCard/crdConfDetailOpen.asp', 'DocEntry=' + objEntry + '&pop=Y&AddPath=', 'objDetails');
			return;
		case 'A1':
			Start('');
			doMyLink('addItem/itmConfDetail.asp', 'DocEntry=' + objEntry + '&pop=Y&AddPath=', 'objDetails');
			return;
		case 'R2':
			Start('');
			doMyLink('cxcRctDetailOpen.asp', 'DocType=-2&DocEntry=' + objEntry + '&pop=Y&AddPath=', 'objDetails');
			return;	
		case 'D3':
			Start('');
			doMyLink('cxcDocDetailOpen.asp', 'DocType=-2&DocEntry=' + objEntry + '&pop=Y&AddPath=', 'objDetails');
			return;	
	}
}
function checkProcess(id)
{
	$.post("executeConfFetch.asp?d=" + (new Date()).toString(), { Type: 'C', ID: id },
	   function(data){
	     doCheckProcess(data, id);
   });
}
function doCheckProcess(data, id)
{
	if (data.indexOf('{S}') == -1) return;
	
	var result = data.split('{S}');
	var status = result[0];
	var errCode = result[1];
	var errMsg = result[2];
	var objCode = result[3];
	var poolCount = result[4];

	switch (status)
	{
		case 'C':
			document.getElementById('txtStatusProc' + id).innerText = txtTransactionPool.replace('{0}', poolCount);
			setTimeout('checkProcess(' + id + ');', 2000);
			break;
		case 'S':
			document.getElementById('txtStatus' + id).innerText = txtConfirmed;
			processRow = document.getElementById('trProc' + id);
			setTimeout('rowFadeOut(10)', 200);
			break;
		case 'E':
			document.getElementById('act' + id).style.backgroundColor = '#FFD2A6';
			document.getElementById('txtStatus' + id).style.display = 'none';
			document.getElementById('status' + id).style.display = '';
			document.getElementById('status' + id).selectedIndex = 0;
			document.getElementById('txtStatusProc' + id).innerHTML = '<b>' + txtError + ': ' + errMsg + '</b>';
			document.getElementById('imgProc' + id).style.display = 'none';
			document.getElementById('dbStatus' + id).value = 'E';
			break;
		case 'P':
			document.getElementById('txtStatus' + id).innerText = txtProcesing;
			document.getElementById('txtStatusProc' + id).innerText = txtProcesing;
			setTimeout('checkProcess(' + id + ');', 2000);
			break;
		default:
			alert('Submit control does not recognize the object status');
			break;
	}
}
function executeAut(id, flowID, lineID, status)
{
	var msg = status == 'A' ? txtConfAut : txtConfReject;
	if (!confirm(msg)) return;
	
	var note = '';
	
	$.post('executeAutFetch.asp?d=' + (new Date()).toString(), { Type: 'S', ID: id, FlowID: flowID, LineID: lineID, Note: note, Status: status },
		function(pData)
		{
			var arrData = pData.split('{S}');
			var autID = id + '_' + flowID + '_' + lineID;
			if(arrData[1] == 'Y')
			{
				alert(txtOtherProc);
				document.getElementById('btnApprove' + autID).disabled = true;
				document.getElementById('btnReject' + autID).disabled = true;
			}
			else
			{
				if (status == 'A')
				{
					document.getElementById('btnApprove' + autID).disabled = true;
					document.getElementById('btnReject' + autID).disabled = true;
				}
				else
				{
					var autAutID = document.frmAut.AutID;
					var autLineID = document.frmAut.LineID;
					var autFlowID = document.frmAut.FlowID;
					if (autAutID.length)
					{
						for (var i = 0;i<autAutID .length;i++)
						{
							var chkID = autAutID[i].value;
							if (chkID)
							{
								var chkID = autAutID[i].value + '_' + autFlowID[i].value + '_' + autLineID[i].value;
								document.getElementById('btnApprove' + chkID).disabled = true;
								document.getElementById('btnReject' + chkID).disabled = true;
							}
						}
					}
					else
					{
						document.getElementById('btnApprove' + autID).disabled = true;
						document.getElementById('btnReject' + autID).disabled = true;
					}
				}
			}
		});
	
}

var inProgressGlobal = false;
var processRow;
function executeProcess(id)
{
	document.getElementById('btnSubmit' + id).disabled = true;
	$.post('executeConfFetch.asp?d=' + (new Date()).toString(), { Type: 'P', ID: id },
		function(pData)
		{
			if (chkProcessData(pData, id))
			{
				var txtNote = document.getElementById('txtNote' + id);
				txtNote.disabled = true;
				
				var note = txtNote.value;
				var status = document.getElementById('status' + id).value;
				
			
				$.post("executeConfFetch.asp?d=" + (new Date()).toString(), { Type: 'S', ID: id, Note: note, Status: status },
				   function(data){
				     doExecuteProcess(data, id);
			   });
			}
			else
			{
				document.getElementById('btnSubmit' + id).disabled = false;
			}
		});
	
}
function doExecuteProcess(data, id)
{
	if (data.indexOf('{S}') == -1) return;
	
	var status = parseInt(document.getElementById('status' + id).value);
	
	processRow = document.getElementById('trProc' + id);
	
	var arrData = data.split('{S}');
	
	if (arrData[0] == 'ok')
	{
		document.getElementById('status' + id).style.display = 'none';
		switch (status)
		{
			case 2:
				document.getElementById('trNote' + id).style.display = 'none';
				document.getElementById('trSubmit' + id).style.display = 'none';
				document.getElementById('txtStatus' + id).style.display = '';
				document.getElementById('trProcStatus' + id).style.display = '';
				document.getElementById('act' + id).style.backgroundColor = '#CCFF99';
				document.getElementById('trProc' + id).style.backgroundColor = '#FFFFCC';
				document.getElementById('txtStatusProc' + id).innerText = txtTransactionPool.replace('{0}', arrData[1]);
				document.getElementById('imgProc' + id).style.display = '';
				document.getElementById('dbStatus' + id).value = 'P';
				setTimeout('checkProcess(' + id + ');', 2000);
				break;
			case 3:
				document.getElementById('act' + id).style.backgroundColor = '';
				document.getElementById('txtStatus' + id).innerText = txtRejected;
				document.getElementById('txtStatus' + id).style.display = '';
				document.getElementById('dbStatus' + id).value = 'R';
				setTimeout('rowFadeOut(10)', 200);
				break;
			case 4:
				document.getElementById('act' + id).style.backgroundColor = '';
				document.getElementById('txtStatus' + id).innerText = txtCanceled;
				document.getElementById('txtStatus' + id).style.display = '';
				document.getElementById('dbStatus' + id).value = 'C';
				setTimeout('rowFadeOut(10)', 200);
				break;
		}
	}
}
function showProcess(status, id)
{
	$.post('executeConfFetch.asp?d=' + (new Date()).toString(), { Type: 'P', ID: id },
		function(pData)
		{
			if (chkProcessData(pData, id))
			{
			    processRow = document.getElementById('trProc' + id);
				if (parseInt(status) > 1)
				{
					document.getElementById('trNote' + id).style.display = '';
					document.getElementById('txtNote' + id).disabled = false;
					document.getElementById('trSubmit' + id).style.display = '';
					document.getElementById('btnSubmit' + id).disabled = false;
					if (document.getElementById('dbStatus' + id).value != 'E')
					    setTimeout('rowFadeIn(1, ' + id + ')', 200); // initial pause
					else
					{
						document.getElementById('txtNote' + id).focus();
						$.post("executeConfFetch.asp?d=" + (new Date()).toString(), { Type: 'N', ID: id },
						   function(data){
						     doLoadNote(data, id);
					   });
					}
				}
				else
				{
					if (document.getElementById('dbStatus' + id).value != 'E')
					    setTimeout('rowFadeOut(10)', 200); // initial pause
					else
					{
						document.getElementById('trNote' + id).style.display = 'none';
						document.getElementById('trSubmit' + id).style.display = 'none';
					}
				}
			}
		});
}
function chkProcessData(data, id)
{
	var arrData = data.split('{S}');
	if (arrData[0] == 'O' || arrData[0] == 'E')
	{
		return true;
	}
	else
	{
		document.getElementById('trNote' + id).style.display = 'none';
		document.getElementById('trSubmit' + id).style.display = 'none';
		document.getElementById('status' + id).style.display = 'none';
		document.getElementById('trProc' + id).style.display = '';
		document.getElementById('trProc' + id).style.backgroundColor = '#FFD2A6';
		document.getElementById('trProcStatus' + id).style.display = '';
		document.getElementById('act' + id).style.backgroundColor = '';
		document.getElementById('txtStatus' + id).style.display = '';
		document.getElementById('imgProc' + id).style.display = 'none';
		switch (arrData[0])
		{
			case 'S':
				document.getElementById('txtStatus' + id).innerText = txtConfirmed;
				document.getElementById('dbStatus' + id).value = 'S';
				break;
			case 'P':
				document.getElementById('txtStatus' + id).innerText = txtProcesing;
				document.getElementById('dbStatus' + id).value = 'P';
				break;
			case 'R':
				document.getElementById('txtStatus' + id).innerText = txtRejected;
				document.getElementById('dbStatus' + id).value = 'R';
				break;
			case 'C':
				document.getElementById('txtStatus' + id).innerText = txtCanceled;
				document.getElementById('dbStatus' + id).value = 'C';
				break;
		}
		var endMsg = msgAllreadyProc.replace('{0}', arrData[3]).replace('{1}', arrData[1]).replace('{2}', arrData[2]);
		if (arrData[4] != '') endMsg += '\n{0}:{1}'.replace('{0}', txtNote).replace('{1}', arrData[4])
		document.getElementById('txtStatusProc' + id).innerText = endMsg;
		return false;
	}
}
function doLoadNote(data, id)
{
	document.getElementById('txtNote' + id).value = data;
}
function rowFadeIn(countUp, id)
{
    processRow.style.display = '';
	for (var i=0; i<processRow.cells.length; i++) 
	{
		processRow.cells[i].style.filter = 'alpha(opacity=' + (countUp * 10) + ')'; // IE
	}
	processRow.style.opacity = countUp / 10; // CSS 3
	countUp += 1;
	if (countUp < 10) 
	{
		setTimeout('rowFadeIn(' + countUp + ', ' + id + ')', 75); // remaining pauses
	} 
	else 
	{
		inProgressGlobal = false;
	    document.getElementById('txtNote' + id).focus();
	}
}
function rowFadeOut(countUp)
{
	for (var i=0; i<processRow.cells.length; i++) 
	{
		processRow.cells[i].style.filter = 'alpha(opacity=' + (countUp * 10) + ')'; // IE
	}
	processRow.style.opacity = countUp / 10; // CSS 3
	countUp -= 1;
	if (countUp > 1)
	{
		setTimeout('rowFadeOut(' + countUp + ')', 75); // remaining pauses
	}
	else
	{
		inProgressGlobal = false;
	    processRow.style.display = 'none';
	}
}
