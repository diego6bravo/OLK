var SaveImgField;
var SaveImgImage;
var SaveImgMaxSize;

function setPayFlow(logNum)
{
	setFlowAlertVars('R2', '', 'document.Form1.submit();', 'payments/submit.asp?submitCmd=update&Confirm=Y');
	flowDraftFld = 'document.Form1.Draft.value';
}

function changeCur()
{
	var selCur = document.getElementById('DocCur').value;
	
	$.post("payments/paymentProcess.asp?d=" + (new Date()).toString(), { Field: 'DocCur', FieldType: 'S', Value: selCur },
	   function(data){
	   		/*
	   		var arrData = data.split('{S}');
	   		curRate = arrData[0]
	   		
			var DocCurRate = document.getElementById('DocCurRate');
			DocCurRate.style.display = MainCur != selCur ? '' : 'none';
			DocCurRate.value = curRate;
			*/
			window.location.reload();
   });
}

function doProc(fld, fldType, value)
{
	$.post("payments/paymentProcess.asp?d=" + (new Date()).toString(), { Field: fld, FieldType: fldType, Value: value },
	   function(data){
	     if (data != 'ok')
	     {
	     	alert(txtErrSaveData);
	     }
   });
}

function doGetTotal()
{
	$.post("payments/paymentProcess.asp?d=" + (new Date()).toString(), { Field: 'GetDocTotal' },
	   function(data){
	   	document.Form1.importe.value = DocCur + ' ' + data;
   });
}

function doProcLine(fld, fldType, value, docType, docNum, instID)
{
	$.post("payments/paymentProcess.asp?d=" + (new Date()).toString(), { Field: fld, FieldType: fldType, Value: value, DocType: docType, DocNum: docNum, InstID: instID },
	   function(data)
	   {
	     if (data != 'ok')
	     {
	     	alert(txtErrSaveData);
	     }
	     else
	     {
	     	if (fld == 'Check' && value == 'Y') return;
		    doGetTotal();
		 }
   });
}


function getImg(Field, Img, MaxSize)
{
	SaveImgField = Field;
	SaveImgImage = Img;
	SaveImgMaxSize = MaxSize;
	OpenWin = this.open('../upload/fileupload.aspx?ID=' + dbID + '&style=../design/' + selDes + '/style/stylePopUp.css', "ImagePicker", "toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no, width=300,height=111");
}
function changepic(img_src) {
SaveImgField.value = img_src;
SaveImgImage.src = "../pic.aspx?filename=" + img_src + "&MaxSize=" + SaveImgMaxSize +'&dbName=' + dbName;
doProc(SaveImgField.name, 'S', img_src);
}

var objField
function datePicker(page, w, h, s, r, o) {
objField = o
OpenWin = this.open(page, "datePicker", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable="+r+", width="+w+",height="+h);
OpenWin.focus()
}
function setTimeStamp2(Action, varDate) { 
objField.value = varDate }

function getCheckVal() {
var checkvalue = ''
for (i=0, n=document.Form1.sourceDocType.length; i<n; i++) {
   if (document.Form1.sourceDocType[i].checked) {
      var checkvalue = document.Form1.sourceDocType[i].value;
      break;
   }
}
return checkvalue
}
function setTimeStamp(retAction, varDate) { retAcct.value = varDate }

function doCheckDel(Img, LogNum)
{
	if (!document.getElementById('countRow' + LogNum).checked)
	{
		document.getElementById('countRow' + LogNum).checked = true;
		Img.src = 'images/checkbox_on.jpg';
	}
	else
	{
		document.getElementById('countRow' + LogNum).checked = false;
		Img.src = 'images/checkbox_off.jpg';
	}
}
function Pay(page, retAction) {
retAcct = retAction
OpenWin = this.open(page, "DatePicker", "toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no, width=240,height=220");
OpenWin.focus()
}


function checkbox(img, docType, docNum, instID, chk, pay, saldo, Cur)
{
	doCheckDel(img, docType + '_' + docNum + '_' + instID);
	if (chk.checked)
	{
		pay.disabled = false;
		pay.className= 'input';
		//doProcLine('Check', 'S', 'Y', docType, docNum, instID);
		doProcLine('SumApplied', 'N', saldo, docType, docNum, instID);
	}
	else
	{
		pay.disabled = true;
		pay.className = 'InputDes';
		doProcLine('Check', 'S', 'N', docType, docNum, instID);
	}
}

function chkThis(Field, FType, EditType, FSize, Restore)
{
	if (Field.value == '' && Restore != null) Field.value = Restore;
	
	switch (FType)
	{
		case "A":
			if (Field.value.length > FSize)
			{
				alert(txtValFldMaxChar.replace("{0}", FSize));
				Field.value = Field.value.substring(0, FSize);
			}
			break;
		case "N":
			switch (EditType)
			{
				case " ":
					if (!MyIsNumeric(getNumeric(Field.value)) && Field.value != "")
					{
						Field.value = "";
						alert(txtValNumVal);
					}
					else if(pasreFloat(getNumeric(Field.value))-parseInt(getNumeric(Field.value)) != 0)
					{
						Field.value = "";
						alert(txtValNumValWhole);
					}
					break
				case "T":
					break
			}
			break;
		case "B":
			if (!MyIsNumeric(getNumeric(Field.value)) && Field.value != "")
			{
				alert(txtValNumVal);
				Field.value = Restore == null ? '' : Restore;
			}
			else
			{
				switch (EditType)
				{
					case "R":
						Field.value = OLKFormatNumber(parseFloat(getNumeric(Field.value)), RateDec);
						break;
					case "S":
						Field.value = OLKFormatNumber(parseFloat(getNumeric(Field.value)), SumDec);
						break;
					case "P":
						Field.value = OLKFormatNumber(parseFloat(getNumeric(Field.value)), PriceDec);
						break;
					case "Q":
						Field.value = OLKFormatNumber(parseFloat(getNumeric(Field.value)), QtyDec);
						break;
					case "%":
						Field.value = OLKFormatNumber(parseFloat(getNumeric(Field.value)), PercentDec);
						break;
					case "M":
						Field.value = OLKFormatNumber(parseFloat(getNumeric(Field.value)), MeasureDec);
						break;
				}
			}
			break;
	}
}

var noMsg = false;
function goAdd(Confirm, DocConf, Draft, Authorize)
{
	document.Form1.Confirm.value = Confirm;
	document.Form1.DocConf.value = DocConf;
	document.Form1.Draft.value = Draft;
	noMsg = true;
	document.Form1.finish.value = "Y";
	document.Form1.B2.click()
	;
}


function updatePagado(pagVal)
{
	document.Form1.pagado.value = DocCur + " " + pagVal;
}

function saldofuera(chk, val)
{
	var SumApplied = parseFloat(getNumeric(document.Form1.pagado.value).replace(DocCur, ''));
	if (chk.checked)
	{
		updatePagado(OLKFormatNumber(SumApplied+val,SumDec));
		doProc("SaldoFuera", "S", "Y");
	}
	else
	{
		updatePagado(OLKFormatNumber(SumApplied-val,SumDec));
		doProc("SaldoFuera", "S", "Y");
	}
}
