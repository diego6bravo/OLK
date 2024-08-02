function chkNewMsg()
{
	var curDate = new Date();
	var url='messages/messageAlertData.asp?date=' + curDate.getTime();

	xmlHttp=GetXmlHttpObject(setNewMsg);
	xmlHttp.open("GET", url , true);
	xmlHttp.send(null);
}

function setNewMsg()
{
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
	{
		var arrMsgStr = xmlHttp.responseText;
	
		if (arrMsgStr != '')
		{
			var msgCount = arrMsgStr.split('{C}')[0];
			
			arrMsg = arrMsgStr.split('{C}')[1].split('{R}');
		
			var tbl = document.getElementById('tblAlert');
			var imgUrgent = '<img alt="" src="images/mail_icon_urgent_small.gif">';
			var pop = false;
			
			for (var i=tbl.rows.length-1;i>=0;i--)
			{
				tbl.deleteRow(i);
			}
			
			if (arrMsgStr.indexOf('{S}') != -1)
			{
				for (var i = 0;i<arrMsg.length;i++)
				{
					var arrMsgVals = arrMsg[i].split('{S}');
					var msgId = arrMsgVals[0];
					var msgSubject = arrMsgVals[1];
					var msgUrgent = arrMsgVals[2];
					var msgStatus = arrMsgVals[3];
					var msgTime = arrMsgVals[4];
					var msgDate = arrMsgVals[5];
					var isNew = arrMsgVals[6] == 'Y';
					if (arrMsgVals[7] == 'Y') pop = true;
					var msgType = arrMsgVals[8];
					var cardType = arrMsgVals[9];
					
					var lastRow = tbl.rows.length;
					var row = tbl.insertRow(lastRow);
					row.style.cursor = 'hand';
					row.id = 'alrMsg' + msgId;
					
					if (!isNew)
					{
						row.className = 'TablasNoticias';
						row.onmouseover = function(){this.className='hlt'};
						row.onmouseout = function(){this.className = 'TablasNoticias'};
					}
					else
					{
						row.className = 'MsgAlert';
						row.onmouseover = function(){this.className='MsgAlertHlt'};
						row.onmouseout = function(){this.className = 'MsgAlert'};
					}
					row.onclick = function(){alertViewMsg(this.id)};
					
					var cell0 = row.insertCell(0);
					cell0.style.width = '20px';
					cell0.style.textAlign = 'center';
					cell0.style.verticalAlign = 'top';
					cell0.innerHTML = getMsgTypeImg(msgType, cardType);
					
					var cell1 = row.insertCell(1);
					cell1.style.width = '10px';
					cell1.style.textAlign = 'center';
					cell1.style.verticalAlign='top';
					if (msgUrgent == 'Y') cell1.innerHTML = imgUrgent;
					
					var cell2 = row.insertCell(2);
					cell2.style.width = '130px';
					cell2.innerHTML = msgDate + ' ' + msgTime;
					
					var cell3 = row.insertCell(3);
					cell3.innerHTML = msgSubject;
				}
			}

			var lastRow = tbl.rows.length;
			var row = tbl.insertRow(lastRow);
			row.className = 'TablasNoticias';
			row.style.cursor = 'hand';
			row.id = 'alrMsg' + msgId;
			row.onmouseover = function(){this.className='hlt'};
			row.onmouseout = function(){this.className = 'TablasNoticias'};
			row.onclick = function(){window.location.href='?cmd=home&onlyMsg=Y'};
			
			var cell1 = row.insertCell(0);
			cell1.style.textAlign = 'center';
			cell1.colSpan = 4;
			cell1.verticalAlign = 'top';
			cell1.innerHTML = '<img src="design/0/images/buzon_cion.gif">&nbsp;' + txtGo2Inbox;
			
			if (parseInt(msgCount) > 1)
			{
				document.getElementById('txtMsgAlert').innerHTML = txtNewMsgs.replace('{0}', msgCount);
			}
			else if (parseInt(msgCount) == 1)
			{
				document.getElementById('txtMsgAlert').innerHTML = txtNewMsg;
			}
			else
			{
				document.getElementById('txtMsgAlert').innerHTML = txtNewMsgs.replace('{0}', 0);
			}
		
			if (!pop && alertOpened == null) 
			{
				document.getElementById('divAlertMsg').style.display = 'none';
				alertOpened = false;
				setAlertPos();
			}
			else if (pop)
			{
				if (alertOpened == null)
				{
					document.getElementById('divAlertMsg').style.display = '';
					document.getElementById('imgAlertOpen').src = 'images/arrow_down_white.gif';
					alertOpened = true;
					setAlertPos();
				}
				else if (!alertOpened)
				{
					openMsgAlert();
				}
			}
			
			if (!alertLoaded)
			{
				document.getElementById('divAlert').style.display = '';
			}
			
			setTimeout('chkNewMsg();', 10000);
			
			alertLoaded = true;
		}
	}
}

function getMsgTypeImg(msgType, cardType)
{
	var imgStr = '';
	var altStr = '';
	
	switch (msgType)
	{
		case 'S':
			imgStr = 'alert';
			altStr = DtxtAlert;
			break;
		case 'C':
			switch (cardType)
			{
				case 'C':
					imgStr = 'supplier';
					altStr = txtClient;
					break;
				case 'S':
					imgStr = 'client';
					altStr = DtxtSupplier;
					break;
				case 'L':
					imgStr = 'lead';
					altStr = DtxtLead;
					break;
			}
			break;
		case 'V':
			imgStr = 'agent';
			altStr = txtAgent;
			break;
		case 'B':
			imgStr = 'system';
			altStr = DtxtSystem;
			break;
		case 'E':
			imgStr = 'alert_red';
			altStr = DtxtError;
			break;
	}

	return '<img border="0" src="ventas/images/icon_' + imgStr + '.gif" alt="' + altStr + '">';
}

function openMsgAlert()
{
	if (!alertOpened)
	{
		document.getElementById('divAlertMsg').style.display = '';
		document.getElementById('imgAlertOpen').src = 'images/arrow_down_white.gif';
		alertOpened = true;
	}
	else
	{
		alertOpened = false;
	}
	setAlertPos(true);
}

var alertSetTop;
var alertOpened;
var alertLoaded = false;
function alertPosEffect()
{
	cTop = parseInt(document.getElementById('divAlert').style.top.replace('px', ''));
	if (cTop > alertSetTop)
	{
		cTop -= 5; //Abrir
		if (cTop <= alertSetTop)
		{
			cTop = alertSetTop;
		}
		else
		{
			setTimeout('alertPosEffect();', 5);
		}
		document.getElementById('divAlert').style.top = cTop + 'px';
	}
	else
	{
		cTop += 5; //Cerrar
		if (cTop >= alertSetTop)
		{
			cTop = alertSetTop;
			document.getElementById('divAlertMsg').style.display = 'none';
			document.getElementById('imgAlertOpen').src = 'images/arrow_up_white.gif';
			document.getElementById('divAlert').style.width = '110px';
			var vLeft = 0;
			if (rtl != '') vLeft += 6;
			document.getElementById('divAlert').style.left = vLeft + 'px';
		}
		else
		{
			setTimeout('alertPosEffect();', 5);
		}
		document.getElementById('divAlert').style.top = cTop + 'px';
	}
}

function setAlertPos()
{
	setAlertPos(false);
}

function setAlertPos(effect)
{
	if (alertOpened) document.getElementById('scrollAlert3').scrollTop = 0;
	
	var vTop = document.body.offsetHeight-173;
	var vLeft = 0;
	
	if (alertOpened) 
	{
		vTop -= 95;
		document.getElementById('divAlert').style.width = (document.body.offsetWidth-30) + 'px';
	}
	else
	{
		if (!effect) document.getElementById('divAlert').style.width = '110px';
	}
	
	if (rtl != '')
	{
		if (alertOpened)
		{
			vLeft -= document.body.offsetWidth-145;
		}
		else
		{
			vLeft += 6;
		}
	}
	
	if (!alertOpened && rtl != '' && !effect || alertOpened) document.getElementById('divAlert').style.left = vLeft + 'px';

	if (effect)
	{
		alertSetTop = vTop;
		setTimeout('alertPosEffect();', 25);
	}
	else
	{
		document.getElementById('divAlert').style.top = vTop + 'px';
	}
}

function alertViewMsg(OlkLog)
{
	doMyLink('messagedetail.asp', 'olklog=' + OlkLog.replace('alrMsg', ''), '');
}

chkNewMsg();
