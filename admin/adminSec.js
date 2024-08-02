function reloadEdit()
{
	document.frmAddEditSec.action = 'adminSecEdit.asp';
	document.frmAddEditSec.submit();
}
function valSecFrm()
{
	SecName = document.frmSec.SecName;
	for (var i = 0;i<SecName.length;i++)
	{
		if (SecName(i).value == '')
		{
			alert(txtValSecNam);
			SecName(i).focus();
			return false;
		}
	}
	return true;
}
function valFrmEdit()
{
	var frm = document.frmAddEditSec;
	
	if (frm.NewName.value == '')
	{
		alert(txtValSecNam);
		frm.NewName.focus();
		return false;
	}
	
	switch (frm.Type.value)
	{
		case 'L':
			if (frm.NewLink.value == '')
			{
				alert(txtValLink);
				frm.NewLink.focus();
				return false;
			}
			break;
		case 'R':
			if (frm.rsIndex.selectedIndex == 0)
			{
				alert(txtValRep);
				frm.rsIndex.focus();
				return false;
			}
	}
	
	document.frmAddEditSec.action = 'adminSubmit.asp';
	return true;
}
function chkSec(chk)
{
	secName = document.getElementById('SecName' + Right(chk.name, chk.name.length-6)).value;
	var errMsg = '';

	if (chk.checked)
	{
		switch (chk.name)
		{
			case 'StatusS2':
				if (!document.getElementById('StatusS-4').checked) errMsg += '\n' + document.getElementById('SecNameS-4').value;
				if (!document.getElementById('StatusS-3').checked) errMsg += '\n' + document.getElementById('SecNameS-3').value;
				if (!document.getElementById('StatusS-2').checked) errMsg += '\n' + document.getElementById('SecNameS-2').value;
				break;
			case 'StatusS1':
				if (!document.getElementById('StatusS-3').checked) errMsg += '\n' + document.getElementById('SecNameS-3').value;
				if (!document.getElementById('StatusS-2').checked) errMsg += '\n' + document.getElementById('SecNameS-2').value;
				break;
			case 'StatusS0':
				if (!document.getElementById('StatusS-3').checked) errMsg += '\n' + document.getElementById('SecNameS-3').value;
				break;
			case 'StatusS-2':
				if (!document.getElementById('StatusS-3').checked) errMsg += '\n' + document.getElementById('SecNameS-3').value;
				break;
		}
	}
	else
	{
		switch (chk.name)
		{
			case 'StatusS-4':
				if (document.getElementById('StatusS2').checked) errMsg += '\n' + document.getElementById('SecNameS2').value;
				break;
			case 'StatusS-3':
				if (document.getElementById('StatusS2').checked) errMsg += '\n' + document.getElementById('SecNameS2').value;
				if (document.getElementById('StatusS1').checked) errMsg += '\n' + document.getElementById('SecNameS1').value;
				if (document.getElementById('StatusS0').checked) errMsg += '\n' + document.getElementById('SecNameS0').value;
				if (document.getElementById('StatusS-2').checked) errMsg += '\n' + document.getElementById('SecNameS-2').value;
				break;
			case 'StatusS-2':
				if (document.getElementById('StatusS2').checked) errMsg += '\n' + document.getElementById('SecNameS2').value;
				if (document.getElementById('StatusS1').checked) errMsg += '\n' + document.getElementById('SecNameS1').value;
				break;
		}
	}
	
	if (errMsg != '')
	{
		if (chk.checked)
		{
				if(confirm(txtValActiveDep.replace('{0}', secName) + '\n' + 
			errMsg + '\n\n' + txtValActiveDepEnd))
			{
				switch (chk.name)
				{
					case 'StatusS2':
						document.getElementById('StatusS-4').checked = true;
						document.getElementById('StatusS-3').checked = true;
						document.getElementById('StatusS-2').checked = true;
						break;
					case 'StatusS1':
						document.getElementById('StatusS-3').checked = true;
						document.getElementById('StatusS-2').checked = true;
						break;
					case 'StatusS0':
						document.getElementById('StatusS-3').checked = true;
						break;
					case 'StatusS-2':
						document.getElementById('StatusS-3').checked = true;
						break;
				}
			}
			else { chk.checked = false; }
		}
		else
		{
				if(confirm(txtValDeactiveSec.replace('{0}', secName) + '\n' + 
			errMsg + '\n\n' + txtValDeactiveSecNow))
			{
				switch (chk.name)
				{
					case 'StatusS-4':
						document.getElementById('StatusS2').checked = false;
						break;
					case 'StatusS-3':
						document.getElementById('StatusS2').checked = false;
						document.getElementById('StatusS1').checked = false;
						document.getElementById('StatusS0').checked = false;
						document.getElementById('StatusS-2').checked = false;
						break;
					case 'StatusS-2':
						document.getElementById('StatusS2').checked = false;
						document.getElementById('StatusS1').checked = false;
						break;
				}
			}
			else { chk.checked = true; }
		}
	}
}
