function delSec(secID)
{
	$('#trSec' + secID).remove();
}
function doAdd()
{
	var cmbSec = document.getElementById('cmbSec');
	
	var SecID = cmbSec.value;
	if (SecID == '') return;
	
	var secName = cmbSec.options.item(cmbSec.selectedIndex).innerText;
	cmbSec.options.remove(cmbSec.selectedIndex);
	cmbSec.selectedIndex = 0;
	
	var NewOrderID = document.getElementById('NewOrderID').value;
	document.getElementById('NewOrderID').value = parseInt(NewOrderID)+1;
	
	var strRow = '';
	strRow += '<tr class="TblRepTbl" id="trSec' + SecID + '">' ;
	strRow += '<td>' + secName + '<input type="hidden" name="SecID" value="' + SecID + '"></td>' ;
	strRow += '<td>' ;
	strRow += '<table cellpadding="0" cellspacing="0" border="0" width="80">' ;
	strRow += '<tr>' ;
	strRow += '<td>' ;
	strRow += '<input type="text" name="OrderID' + SecID + '" id="OrderID' + SecID + '" size="7" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="' + NewOrderID + '">' ;
	strRow += '</td>' ;
	strRow += '<td valign="middle">' ;
	strRow += '<table cellpadding="0" cellspacing="0" border="0">' ;
	strRow += '<tr>' ;
	strRow += '<td><img src="images/img_nud_up.gif" id="btnOrderID' + SecID + 'Up"></td>' ;
	strRow += '</tr>' ;
	strRow += '<tr>' ;
	strRow += '<td><img src="images/spacer.gif"></td>' ;
	strRow += '</tr>' ;
	strRow += '<tr>' ;
	strRow += '<td><img src="images/img_nud_down.gif" id="btnOrderID' + SecID + 'Down"></td>' ;
	strRow += '</tr>' ;
	strRow += '</table>' ;
	strRow += '</td>' ;
	strRow += '</tr>' ;
	strRow += '</table>' 
	strRow += '</td>' ;
	strRow += '<td style="width: 16px">' ;
	strRow += '<img border="0" src="images/remove.gif" onclick="javascript:if(confirm(\'|L:txtConfDelSec|\'.replace(\'{0}\', \'' + escape(secName) + '\')))delSec(' + SecID + ');"></td>' ;
	strRow += '</tr>' ;

	$("#tbSec").append(strRow);
	
	NumUDAttach('frmGroupEdit', 'OrderID' + SecID, 'btnOrderID' + SecID + 'Up', 'btnOrderID' + SecID + 'Down');

}


function valFrm()
{
	if (document.frmGroups.GroupID)
	{
		if (document.frmGroups.GroupID.length)
		{
			for (var i = 0;i<document.frmGroups.GroupID.length;i++)
			{
				var txt = document.getElementById('GroupName' + document.frmGroups.GroupID[i].value);
				if (txt.value == '')
				{
					alert(txtValGrmNam);
					txt.focus();
					return false;
				}	
			}
		}
		else
		{
			var txt = document.getElementById('GroupName' + document.frmGroups.GroupID.value);
			if (txt.value == '')
			{
				alert(txtValGrmNam);
				txt.focus();
				return false;

			}
		}
	}
	return true;
}