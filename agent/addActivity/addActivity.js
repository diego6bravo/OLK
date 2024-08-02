var SaveImgField;
var SaveImgImage;
var SaveImgMaxSize;
var objFieldType;

function setActFlow()
{
	setFlowAlertVars('C2', '', 'document.frmAddActivity.DocConf.value = typeIDs;document.frmAddActivity.submit();', '');
}

function getImg(Field, Img, MaxSize)
{
	SaveImgField = Field;
	SaveImgImage = Img;
	SaveImgMaxSize = MaxSize;
	Start('upload/fileupload.aspx?ID=' + dbID + '&style=../design/' + SelDes + '/style/stylePopUp.css',300,100,'no')
}

function Start(page, w, h, s) {
OpenWin = this.open(page, "ImagePicker", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable=no, width="+w+",height="+h);
}

function changepic(value) {
doProc(SaveImgField.name, 'S', value);
SaveImgField.value = value;
SaveImgImage.src = "pic.aspx?filename=" + value + "&MaxSize=" + SaveImgMaxSize + '&dbName=' + dbName;
}
function showUDF(id)
{
	if (id == -1) id = '_1';
	
	var tdShowUDF = document.getElementById('tdShowUDF' + id);
	var trUDF = document.getElementById('trUDF' + id);

	trUDF.style.display = '';
	tdShowUDF.innerHTML = '[-]';
}


Calendar.setup({
    inputField     :    "Recontact",     // id of the input field
    ifFormat       :    CalendarFormat,      // format of the input field
    button         :    "btnBeginDate",  // trigger for the calendar (button ID)
    align          :    "Bl",           // alignment (defaults to "Bl")
    singleClick    :    true
});

Calendar.setup({
    inputField     :    "endDate",     // id of the input field
    ifFormat       :    CalendarFormat,      // format of the input field
    button         :    "btnEndDate",  // trigger for the calendar (button ID)
    align          :    "Bl",           // alignment (defaults to "Bl")
    singleClick    :    true
});

var objField;
function datePicker(page, w, h, s, r, o, pType) 
{
	objField = o;
	objFieldType = pType;
	OpenWin = this.open(page, "datePicker", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable="+r+", width="+w+",height="+h);
	OpenWin.focus();
}
function doProc(fld, fldType, value)
{
	$.post("addActivity/addActivityProcess.asp?d=" + (new Date()).toString(), { Field: fld, FieldType: fldType, Value: value },
	   function(data){
	     if (data != 'ok')
	     {
	     	alert(txtErrSaveData);
	     }
   });
}
function setTimeStamp(Action, value) 
{ 
	objField.value = value;
	doProc(objField.name, objFieldType, value);
}

function goAdd(Confirm, DocConf)
{
	document.frmAddActivity.Confirm.value = Confirm;
	document.frmAddActivity.DocConf.value = DocConf;
	document.frmAddActivity.btnAdd.click();
}
function changeAction(action)
{
	var display = action == 'M' ? '' : 'none';
	document.getElementById('trAddress1').style.display = display;
	document.getElementById('trAddress2').style.display = display;
	document.getElementById('optTentative').style.display = display;
	
	display = action != 'T' ? '' : 'none';
	document.getElementById('tdENDTime').style.display = display;
	document.getElementById('optReminder').style.display = display;
	document.getElementById('trReminder').style.display = display;
	document.getElementById('tdBeginTime').style.display = display;
	
	display = action != 'T' && action != 'E' ? '' : 'none';
	document.getElementById('txtDuration').style.display = display;
	document.getElementById('tblDuration').style.display = display;
	
	display = action == 'T' ? '' : 'none';
	document.getElementById('txtStatus').style.display = display;
	document.getElementById('Status').style.display = display;
	
	display = action != 'E' ? '' : 'none';
	document.getElementById('txtENDTime').style.display = display;
	document.getElementById('tblENDTime').style.display = display;
	document.getElementById('txtLocation').style.display = display;
	document.getElementById('Location').style.display = display;

	var startText = '';
	var endText = '';
	switch (action)
	{
		case 'T':
			startText = txtStartDate;
			endText = txtDueDate;
			break;
		case 'E':
			startText = txtTime;
			endText = '';
			break;
		default:
			startText = txtStartTime;
			endText = txtEndTime;
			break;
	}
	document.getElementById('txtBeginTime').innerText = startText;
	document.getElementById('txtENDTime').innerText = endText;
	doProc('Action', 'S', action);
}
function changeType(value)
{
	$.post('addActivity/actGetSmallList.asp?d=' + (new Date()).toString(), { Type: 'S', Value: value }, function(data)
		{
			
			var options = '';
		
			if (data != '')
			{
				var arrData = data.split('{S}');	
				for (var i = 0;i<arrData.length;i++)
				{
					var arrCols = arrData[i].split('{C}');
					options += '<option value="' + arrCols[0] + '">' + arrCols[1] + '</option>';
				}
			}
			
			$('#CntctSbjct').find('option').remove().end().append(options);
			doProc('CntctType', 'N', value);
			doProc('CntctSbjct', 'N', '');
		});
}
function changeContact(value)
{
	$.post('addActivity/actGetSmallList.asp?d=' + (new Date()).toString(), { Type: 'CTel', Code: value }, function(data)
		{
			document.frmAddActivity.Tel.value = data;
		});
}
function changeCountry(value)
{
	$.post('addActivity/actGetSmallList.asp?d=' + (new Date()).toString(), { Type: 'Cnt', Code: value }, function(data)
		{
			
			var options = '';
		
			if (data != '')
			{
				var arrData = data.split('{S}');	
				for (var i = 0;i<arrData.length;i++)
				{
					var arrCols = arrData[i].split('{C}');
					options += '<option value="' + arrCols[0] + '">' + arrCols[1] + '</option>';
				}
			}
			
			$('#State').find('option').remove().end().append(options);
			doProc('Country', 'S', value);
			doProc('State', 'S', '');
		});
}

function GetDatePart(value, p)
{
	switch (p)
	{
		case 'Y':
			return Mid(value, DisplayFormat.indexOf('yyyy'), 4);
			break;
		case 'M':
			return Mid(value, DisplayFormat.indexOf('MM'), 2);
			break;
		case 'D':
			return Mid(value, DisplayFormat.indexOf('dd'), 2);
			break;
	}
}
function getActDate(arg)
{
	switch (arg)
	{
		case 'F':
			var bH = document.frmAddActivity.BeginTimeH.value;
			var bM = document.frmAddActivity.BeginTimeM.value;
			var bS = document.frmAddActivity.BeginTimeS.value;
			
			if (bS == 'PM' && parseInt(bH) != 12) bH = parseInt(bH)+12;
			else if (bS == 'AM' && parseInt(bH) == 12) bH = 0;
		
			var strDateFrom = document.frmAddActivity.Recontact.value;
			var retDate = new Date(GetDatePart(strDateFrom, 'Y'),
					parseInt(GetDatePart(strDateFrom, 'M'))-1,
					GetDatePart(strDateFrom, 'D'), bH, bM);
			return retDate;
			break;
		case 'T':
			var eH = document.frmAddActivity.ENDTimeH.value;
			var eM = document.frmAddActivity.ENDTimeM.value;
			var eS = document.frmAddActivity.ENDTimeS.value;
				
			if (eS == 'PM' && parseInt(eH) != 12) eH = parseInt(eH)+12;
			else if (eS == 'AM' && parseInt(eH) == 12) eH = 0;
			
			var strDateTo = document.frmAddActivity.endDate.value;
			var retDate = new Date(GetDatePart(strDateTo, 'Y'),
					parseInt(GetDatePart(strDateTo, 'M'))-1,
					GetDatePart(strDateTo, 'D'), eH, eM);
			return retDate;
			break;
	}
}


function changeTime(Source)
{
			
	switch (Source)
	{
		case 'beginT':
			changeTimeEnd(getActDate('F'));
			//loadDurType();
			break;
		case 'endT':
			//loadDurType();
			break;
		case 'dur':
			if (!MyIsNumeric(document.frmAddActivity.Duration.value))
			{
				alert(txtValNumVal);
				document.frmAddActivity.Duration.value = document.frmAddActivity.DurationUndo.value;
				document.frmAddActivity.Duration.focus();
			}
			else
			{
				document.frmAddActivity.DurationUndo.value = document.frmAddActivity.Duration.value;
				changeTimeEnd(getActDate('F'));
			}
			break;
	}
	
	var Recontact = document.frmAddActivity.Recontact.value
	var BeginTime = document.frmAddActivity.BeginTimeH.value + ',' + document.frmAddActivity.BeginTimeM.value + ',' + document.frmAddActivity.BeginTimeS.value
	var endDate = document.frmAddActivity.endDate.value;
	var ENDTime = document.frmAddActivity.ENDTimeH.value + ',' + document.frmAddActivity.ENDTimeM.value + ',' + document.frmAddActivity.ENDTimeS.value;
	
	if (Source == 'dir')
	{
		doProc('Recontact', 'D', Recontact);
		doProc('BeginTime', 'T', BeginTime);
		doProc('endDate', 'D', endDate);
		doProc('ENDTime', 'T', ENDTime);
		
		doProc('Duration', 'N', document.frmAddActivity.Duration.value);
		doProc('DurType', 'S', document.frmAddActivity.DurType.value);
	}
	else
	{
		
		$.post('addActivity/actGetSmallList.asp?d=' + (new Date()).toString(), { 
							Type: 'Dur',
							Recontact: Recontact,
							BeginTime: BeginTime,
							endDate: endDate,
							ENDTime: ENDTime
						}, function(data)
			{
				var arrData = data.split('{S}');
				document.frmAddActivity.Duration.value = arrData[0];
				document.frmAddActivity.DurationUndo.value = arrData[0];
				document.frmAddActivity.DurType.value = arrData[1];
			});
	}
}


function changeTimeEnd(dFrom)
{
	durType = document.frmAddActivity.DurType.value;
	switch (durType)
	{
		case 'M':
			durType = 'n';
			break;
		case 'H':
			durType = 'h';
			break;
	}
	dTo = dFrom.dateAdd(durType, document.frmAddActivity.Duration.value);
	
	var endDate = DisplayFormat.replace('yyyy', dTo.getFullYear()).replace('MM', Right('0' + (dTo.getMonth()+1), 2)).replace('dd', Right('0' + dTo.getDate(), 2));
	
	document.frmAddActivity.endDate.value = endDate; 
	var eH = dTo.getHours();
	if (parseInt(eH) > 12)
	{
		eH = eH-12;
		document.frmAddActivity.ENDTimeS.value = 'PM';
	}
	else
	{
		document.frmAddActivity.ENDTimeS.value = 'AM';
	}
	document.frmAddActivity.ENDTimeH.value = Right('0' + eH, 2);
	document.frmAddActivity.ENDTimeM.value = Right('0' + dTo.getMinutes(), 2);
	
}

function changeReminder()
{
	if (!MyIsNumeric(document.frmAddActivity.RemQty.value))
	{
		alert(txtValNumVal);
		document.frmAddActivity.RemQty.value = document.frmAddActivity.RemQtyUndo.value;
		document.frmAddActivity.RemQty.focus();
	}
	else
	{
		document.frmAddActivity.RemQtyUndo.value = document.frmAddActivity.RemQty.value;
	}
}

function changeDocNum()
{
	docNum = document.frmAddActivity.DocNum.value;
	if (docNum != '')
	{
		if (docNum.indexOf('*') == -1)
		{
			$.post('addActivity/actGetSmallList.asp?d=' + (new Date()).toString(), { 
								Type: 'Doc',
								DocType: document.frmAddActivity.DocType.value,
								Value: document.frmAddActivity.DocNum.value
							}, function(data)
				{
					var arrData = data.split('|');
					if (arrData[0] == 'ok')
					{
						document.frmAddActivity.DocEntry.value = arrData[1];
						document.frmAddActivity.DocNum.value = arrData[2];
					}
					else
					{
						document.frmAddActivity.DocEntry.value = '';
						document.frmAddActivity.DocNum.value = '';
						launchSelect();
					}
					doProc('DocEntry', 'I', document.frmAddActivity.DocEntry.value);
				});
		}
		else
		{
			launchSelect();
		}
	}
	else
	{
		document.frmAddActivity.DocEntry.value = '';
		doProc('DocEntry', 'I', document.frmAddActivity.DocEntry.value);
	}
}

function changeDocType()
{
	document.frmAddActivity.DocNum.value = '';
	document.frmAddActivity.DocEntry.value = '';
	document.frmAddActivity.DocNum.disabled = document.frmAddActivity.DocType.selectedIndex == 0;
	doProc('DocType', 'N', document.frmAddActivity.DocType.value);
	doProc('DocEntry', 'I', '');
}

function launchSelect(){
	var retVal = window.showModalDialog('topGetDocLink.asp?Type=DocLink&DocType=' + document.frmAddActivity.DocType.value + '&DocNum=' + document.frmAddActivity.DocNum.value,'','dialogWidth:640px;dialogHeight:500px');
	if (retVal != '' && retVal != null)
	{
		document.frmAddActivity.DocEntry.value = retVal.split(',')[0];
		document.frmAddActivity.DocNum.value = retVal.split(',')[1];
	} 
	else 
	{ 
		document.frmAddActivity.DocEntry.value = '';
		document.frmAddActivity.DocNum.value = '';
	}
}

function GetYesNo(value)
{
	return value ? 'Y' : 'N';
}