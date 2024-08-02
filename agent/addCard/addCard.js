var SaveImgField;
var SaveImgImage;
var SaveImgMaxSize;
var objFieldType;

function setCardFlow(logNum)
{
	setFlowAlertVars('C1', '', 'document.frmAddCard.DocConf.value = typeIDs;document.frmAddCard.doSubmitAdd.value=\'Y\';document.frmAddCard.submit();', 'agentClientSubmit.asp?RetVal=' + logNum + '&Confirm=Y');
}

function getImg(Field, Img, MaxSize)
{
	SaveImgField = Field;
	SaveImgImage = Img;
	SaveImgMaxSize = MaxSize;
	Start('upload/fileupload.aspx?ID=' + dbID + '&style=../design/' + selDes + '/style/stylePopUp.css',300,100,'no')
}

function Start(page, w, h, s) 
{
	OpenWin = this.open(page, "ImagePicker", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable=no, width="+w+",height="+h);
}

function changepic(value)
{
	SaveImgField.value = value;
	SaveImgImage.src = "pic.aspx?filename=" + value + "&MaxSize=" + SaveImgMaxSize + '&dbName=' + dbName;
	doProc(SaveImgField.name, 'S', value);
}

function datePicker(page, w, h, s, r, o, pType) 
{
	objField = o;
	objFieldType = pType;
	OpenWin = this.open(page, "datePicker", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable="+r+", width="+w+",height="+h);
	OpenWin.focus()
}
function setTimeStamp(Action, value) 
{ 
	objField.value = value; 
	doProc(objField.name, objFieldType, value);
}

function changeCardType(CardType)
{
	switch (CardType)
	{
		case 'C':
			document.frmAddCard.btnAdd.value = ConfirmC ? DtxtConfirm : DtxtAdd;
			if (document.frmAddCard.prevCardType.value == 'S') 
			{
				loadGroups('C');
				if (document.getElementById('cmbGroupNum')) document.frmAddCard.cmbGroupNum.value = DfCustTerm;	
			}
			break;
		case 'S':
			document.frmAddCard.btnAdd.value = ConfirmS ? DtxtConfirm : DtxtAdd;
			if (document.frmAddCard.prevCardType.value != 'S') 
			{
				loadGroups('S');
				if (document.getElementById('cmbGroupNum')) document.frmAddCard.cmbGroupNum.value = DfVendTerm;
			}
			break;
		case 'L':
			document.frmAddCard.btnAdd.value = ConfirmL ? DtxtConfirm : DtxtAdd;
			if (document.frmAddCard.prevCardType.value == 'S') 
			{
				loadGroups('C');
				if (document.getElementById('cmbGroupNum')) document.frmAddCard.cmbGroupNum.value = DfCustTerm;
			}
			break;
	}
	
	document.frmAddCard.prevCardType.value = CardType;
	doProc('CardType', 'S', CardType);
	if (document.getElementById('cmbGroupNum')) doProc('GroupNum', 'N', document.frmAddCard.cmbGroupNum.value);
}

function loadGroups(GrpCardType)
{
	$.post("addCard/addCardFetch.asp?d=" + (new Date()).toString(), { Type: "crdGroups", Value: GrpCardType },
	   function(data){
			var arrData = data.split('{S}');
			
			var cmbGrp = document.frmAddCard.GroupCode;
			for (var i = cmbGrp.length-1;i>=0;i--) cmbGrp.remove(i);
			
			for (var i = 0;i<arrData.length;i++)
			{
				var arrVals = arrData[i].split('{V}');
				cmbGrp.options[i] = new Option(arrVals[1], arrVals[0]);
			}
			
			doProc('GroupCode', 'N', cmbGrp.value);
   });
}
function changeCmpPrivate(value)
{
	if (lawsSet == 'MX' && document.getElementById('LicTradNum'))
	{
		var maxLength = 12;
		if (value == 'I') maxLength = 13;
	
		document.frmAddCard.LicTradNum.maxLength = maxLength;
		document.frmAddCard.LicTradNum.onkeydown = new Function("return chkMax(event, this, {0});".replace('{0}', maxLength));
		if (document.frmAddCard.LicTradNum.value.length > maxLength)
			document.frmAddCard.LicTradNum.value = document.frmAddCard.LicTradNum.value.substring(0, maxLength);
	}
	doProc('CmpPrivate', 'S', value);
}
function doProc(fld, fldType, value)
{
	$.post("addCard/addCardProcess.asp?d=" + (new Date()).toString(), { Field: fld, FieldType: fldType, Value: value },
	   function(data){
	     if (data != 'ok')
	     {
	     	alert(txtErrSaveData);
	     }
   });
}

function chkBP(value)
{
	$.post("addCard/addCardFetch.asp?d=" + (new Date()).toString(), { Type: "chkCode", Value: value },
	   function(data){
	   	document.getElementById('dvCodeErr').style.display = data == 'Y' ? '' : 'none';
	   	document.frmAddCard.btnAdd.disabled = data == 'Y';
	   	document.frmAddCard.CardCode.style.backgroundColor = data == 'Y' ? '#FFD2A6' : '';
   });
}

function goAdd(Confirm, DocConf)
{
	document.frmAddCard.Confirm.value = Confirm;
	document.frmAddCard.DocConf.value = DocConf;
	document.frmAddCard.btnAdd.click();
}
function copyMailAddress() 
{
	document.frmAddCard.Address.value = document.frmAddCard.MailAddres.value;
	document.frmAddCard.City.value = document.frmAddCard.MailCity.value;
	document.frmAddCard.County.value = document.frmAddCard.MailCounty.value;
	document.frmAddCard.Country.value = document.frmAddCard.MailCountr.value;
	document.frmAddCard.ZipCode.value = document.frmAddCard.MailZipCod.value  
}

function copyAddress() 
{
	document.frmAddCard.MailAddres.value = document.frmAddCard.Address.value;
	document.frmAddCard.MailCity.value = document.frmAddCard.City.value;
	document.frmAddCard.MailCounty.value = document.frmAddCard.County.value;
	document.frmAddCard.MailCountr.value = document.frmAddCard.Country.value;
	document.frmAddCard.MailZipCod.value = document.frmAddCard.ZipCode.value; 
}

function showUDF(id)
{
	if (id == -1) id = '_1';
	
	var tdShowUDF = document.getElementById('tdShowUDF' + id);
	var trUDF = document.getElementById('trUDF' + id);

	trUDF.style.display = '';
	tdShowUDF.innerHTML = '[-]';
}

function valFrm()
{
	var LicTradMaxSize = 0;
	switch (lawsSet)
	{
		case 'MX':
		case 'CR':
		case 'GT':
			LicTradMaxSize = document.frmAddCard.CmpPrivate.value == 'C' ? 12 : 13;
			break;
		case 'CL':
			LicTradMaxSize = 13;
			break;
	}
	
	if (document.frmAddCard.CardCode.value == "")
	{
		alert(txtValCod);
		document.frmAddCard.CardCode.focus();
		return false;
	}
	
	if (document.frmAddCard.GroupCode.value == '')
	{
		alert(txtSelGrp);
		document.frmAddCard.GroupCode.focus();
		return false;
	}
	
	if (document.getElementById('LicTradNum'))
	{
		if ((lawsSet == 'MX' || lawsSet == 'CL' || lawsSet == 'CR' || lawsSet == 'GT') && document.frmAddCard.LicTradNum.value == "")
		{
			alert(txtValRFC);
			document.frmAddCard.LicTradNum.focus();
			return false;
		}
	
		if ((lawsSet == 'MX' || lawsSet == 'CR' || lawsSet == 'GT') && document.frmAddCard.LicTradNum.value.length != LicTradMaxSize)
		{
			alert(txtValRFCLen.replace('{0}', LicTradMaxSize))
			document.frmAddCard.LicTradNum.focus();
			return false;
		}
	}
	
	if (!valUDF()) return false;
	
	return true;
}
