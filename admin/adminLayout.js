function doAdd(groupID, colID)
{
	var id = document.getElementById('cmbAddLine' + colID).value;
	if (id != '')
	{
		var arrID = id.split('|');
		var strType = arrID[0];
		var typeID = arrID[1];
		var rowID = document.getElementById('NewRowID' + colID).value;
		var chkActive = document.getElementById('chkNewActive' + colID).checked ? 'Y' : 'N';

		window.location.href='adminLayoutSubmit.asp?cmd=addLine&GroupID=' + groupID + '&ColID=' + colID + '&Type=' + strType + 
								'&ID=' + typeID + '&RowID=' + rowID + '&Active=' + chkActive;
	}
}