doSelLang();
document.onclick=clearSelectLang;

var OpenWin = null;

function chkWin()
{
	if (OpenWin != null) if (!OpenWin.Closed) { OpenWin.focus(); this.blur(); }
}

function clearWin() { OpenWin = null; }

function GetTopPos(inputObj)
{
	
  var returnValue = inputObj.offsetTop;
  while((inputObj = inputObj.offsetParent) != null){
  	returnValue += inputObj.offsetTop;
  }
  return returnValue;
}

function GetLeftPos(inputObj)
{
  var returnValue = inputObj.offsetLeft;
  while((inputObj = inputObj.offsetParent) != null)returnValue += inputObj.offsetLeft;
  return returnValue;
}

var NewValueTradField;
function doFldTrad(Table, ColumnID, ID, ColumnName, Type, NewValue)
{
	if (NewValue != null) NewValueTradField = NewValue;
	page = '';
	document.frmAdminTrad.Table.value = Table;
	document.frmAdminTrad.ColumnID.value = ColumnID;
	document.frmAdminTrad.ID.value = ID;
	document.frmAdminTrad.ColumnName.value = ColumnName;
	document.frmAdminTrad.Type.value = Type;
	if (NewValue != null) document.frmAdminTrad.NewValue.value = NewValue.value;
	else document.frmAdminTrad.NewValue.value = '';
	document.frmAdminTrad.IsNew.value = (NewValue == null ? 'N' : 'Y');
	switch (Type)
	{
		case 'T':
			w = 400;
			h = 82;
			break;
		case 'M':
			w = 400;
			h = 144;
			break;
		case 'R':
			w = 640;
			h = 480;
			break;
	}
	OpenWin = this.open(page, "CtrlWindow", "toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=no, width=" + w + ",height=" + h);
	document.frmAdminTrad.submit();
}
function setNewFldTrad(NewValue)
{
	NewValueTradField.value = NewValue;
	NewValueTradField = null;
}

function doFldNote(PageID, FieldID, FieldKey, NewValue)
{
	if (NewValue != null) NewValueTradField = NewValue;
	page = '';
	document.frmAdminDefinition.PageID.value = PageID;
	document.frmAdminDefinition.FieldID.value = FieldID;
	document.frmAdminDefinition.FieldKey.value = FieldKey;
	if (NewValue != null) document.frmAdminDefinition.NewValue.value = NewValue.value;
	else document.frmAdminDefinition.NewValue.value = '';
	document.frmAdminDefinition.IsNew.value = (NewValue == null ? 'N' : 'Y');
	OpenWin = this.open(page, 'CtrlWindow', 'toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=no,width=400,height=250');
	document.frmAdminDefinition.submit();
}
