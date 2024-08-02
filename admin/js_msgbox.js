function doMsgBox()
{
	var strMsgTable = '<div style="display: none;" class="MsgBoxBG" id="MsgBox"> ' + 
						'</div> ' + 
						'<div id="TblMsgBox" style="position: absolute;z-index: 102;top: 100px;left: 100px; display: none; "> ' + 
						'<table cellpadding="0" border="0" cellspacing="0" class="TblMsgBox"> ' + 
						'	<tr> ' + 
						'		<td class="MsgBoxTtl" id="tblMessageText"></td> ' + 
						'		<td class="MsgBoxTtl" width="68"><img id="imgMsgBox"></td> ' +
						'	</tr> ' + 
						'	<tr> ' + 
						'		<td class="MsgBoxBody" align="right" colspan="2"> ' + 
						'		<table cellpadding="0" cellspacing="10" border="0"> ' + 
						'			<tr> ' + 
						'				<td><input type="button" id="btnMsgBoxYes" value="Yes" class="OlkBtn" /></td> ' + 
						'				<td id="tdMsgBoxNo"><input type="button" id="btnMsgBoxNo" value="No" class="OlkBtn" /></td> ' + 
						'				<td id="tdMsgBoxCancel"><input type="button" id="btnMsgBoxCancel" value="Cancel" class="OlkBtn" /></td> ' + 
						'			</tr> ' + 
						'		</table> ' + 
						'		</td> ' + 
						'	</tr> ' + 
						'</table> ' + 
						'</div> ';
	document.write(strMsgTable);
	
}

function showMsgBox(MessageText, MessageType, YesClickCommand, NoClickCommand, CancelClickCommand, IconType)
{
	document.getElementById('tblMessageText').innerHTML = MessageText.replace('\n', '<br>');
	
	setMsgBoxCommand(document.getElementById('btnMsgBoxYes'), YesClickCommand);
	setMsgBoxCommand(document.getElementById('btnMsgBoxNo'), NoClickCommand);
	setMsgBoxCommand(document.getElementById('btnMsgBoxCancel'), CancelClickCommand);
	
	switch (MessageType)
	{
		case 0:
			document.getElementById('tdMsgBoxNo').style.display = 'none';
			document.getElementById('tdMsgBoxCancel').style.display = 'none';
			document.getElementById('btnMsgBoxYes').value = 'Ok';
			break;
		case 1:
			document.getElementById('tdMsgBoxNo').style.display = 'none';
			document.getElementById('btnMsgBoxYes').value = 'Ok';
			break;
		case 2:
			document.getElementById('tdMsgBoxNo').style.display = '';
			document.getElementById('btnMsgBoxYes').value = 'Yes';
			break;
	}
	
	document.getElementById('MsgBox').style.display = '';
	document.getElementById('TblMsgBox').style.display = '';
	
	switch (IconType)
	{
		case 1:
			document.getElementById('imgMsgBox').src = 'images/questionIcon.gif';
			break;
		case 2:
			document.getElementById('imgMsgBox').src = 'images/errorIcon.gif';
			break;
		default:
			document.getElementById('imgMsgBox').src = 'images/confirmIcon.gif';
			break;
	}
	
	setMsgBoxPos();
	
	document.getElementById('btnMsgBoxYes').focus();
}
function setMsgBoxPos()
{
	if (document.getElementById('MsgBox'))
	if (document.getElementById('MsgBox').style.display == '')
	{
		if (browserDetect() == 'msie')
		{
			document.getElementById('TblMsgBox').style.top = parseInt(document.body.clientHeight/2)-65+document.body.scrollTop;
			document.getElementById('TblMsgBox').style.left = parseInt(document.body.clientWidth/2)-215+document.body.scrollLeft;
			
			document.getElementById('MsgBox').style.height = document.body.scrollHeight;
			document.getElementById('MsgBox').style.width = document.body.scrollWidth ;

		}
		else
		{
			document.getElementById('TblMsgBox').style.top = (parseInt(window.innerHeight/2)-65+document.body.scrollTop) + 'px';
			document.getElementById('TblMsgBox').style.left = (parseInt(window.innerWidth/2)-215+document.body.scrollLeft) + 'px';
			
			document.getElementById('MsgBox').style.height = document.body.scrollHeight + 'px';
			document.getElementById('MsgBox').style.width = document.body.scrollWidth + 'px';
		}
	}
}
function setMsgBoxCommand(Button, Command)
{
	var cmdStr = 'javascript:';
	if (Command != '') cmdStr += Command + ';';
	cmdStr += 'msgBoxClose();';
	if (browserDetect() == 'firefox')
	{
		Button.onclick = new Function('event',cmdStr);
	}
	else
	{
		Button.onclick = new Function(cmdStr);
	}	
}
function msgBoxClose()
{
	document.getElementById('MsgBox').style.display = 'none';
	document.getElementById('TblMsgBox').style.display = 'none';
}
