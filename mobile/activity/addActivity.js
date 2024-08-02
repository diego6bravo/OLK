
function changeType()
{
	document.frmActivity.action = 'operaciones.asp';
	document.frmActivity.cmd.value = 'activity';
	document.frmActivity.CntctSbjct.selectedIndex = 0;
	document.frmActivity.submit();
}
function changeAction()
{
	document.frmActivity.action = 'operaciones.asp';
	document.frmActivity.cmd.value = 'activity';
	document.frmActivity.submit();
}

function changeTime(Source)
{
	document.frmGeneral.cmd.value = 'activityGeneral';
	document.frmGeneral.action = 'operaciones.asp';
	document.frmGeneral.Source.value = Source;
	document.frmGeneral.submit();
}

function changeDocType()
{
	document.frmGeneral.DocNum.value = '';
	document.frmGeneral.DocEntry.value = '';
	document.frmGeneral.DocNum.disabled = document.frmGeneral.DocType.selectedIndex == 0;
}

function changeReminder()
{
	if (!MyIsNumeric(document.frmGeneral.RemQty.value))
	{
		alert(txtValNumVal);
		document.frmGeneral.RemQty.value = document.frmGeneral.RemQtyUndo.value;
		document.frmGeneral.RemQty.focus();
	}
	else
	{
		document.frmGeneral.RemQtyUndo.value = document.frmGeneral.RemQty.value;
	}
}
function getCalUDF(AliasID)
{
	document.frmUDF.action = 'operaciones.asp';
	document.frmUDF.editVar.value = AliasID;
	document.frmUDF.cmd.value = 'UDFCal';
	document.frmUDF.submit();
}
function getValUDF(AliasID)
{
	document.frmUDF.action = 'operaciones.asp';
	document.frmUDF.editVar.value = AliasID;
	document.frmUDF.cmd.value = 'UDFQry';
	document.frmUDF.submit();
}

function getCal(AliasID)
{
	switch (AliasID)
	{
		case 'Recontact':
			document.frmGeneral.Source.value = 'beginT';
			break;
		case 'endDate':
			document.frmGeneral.Source.value = 'endT';
			break;
	}
	document.frmGeneral.action = 'operaciones.asp';
	document.frmGeneral.editVar.value = AliasID;
	document.frmGeneral.cmd.value = 'UDFCal';
	document.frmGeneral.submit();
}

function changeDocNum()
{
	document.frmGeneral.action = 'operaciones.asp';
	document.frmGeneral.cmd.value = 'activityGeneral';
	document.frmGeneral.editVar.value = 'DocNum';
	document.frmGeneral.submit();
}