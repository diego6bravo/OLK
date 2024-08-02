function chkMax(e, f, m)
{
	if(f.value.length == m && (e.keyCode != 8 && e.keyCode != 9 && e.keyCode != 35 && e.keyCode != 36 && e.keyCode != 37 
	&& e.keyCode != 38 && e.keyCode != 39 && e.keyCode != 40 && e.keyCode != 46 && e.keyCode != 16))return false; else return true;
}
var OpenWin = null;
function chkWin() { if (OpenWin != null) if (!OpenWin.closed) OpenWin.focus() }

function Start(o, page, w, h, s, r) {
objField = o
OpenWin = this.open(page, "queryWin", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable="+r+", width="+w+",height="+h);
OpenWin.focus()
}

function chkNum(fld, dType)
{
	if (dType != 'nvarchar')
	{
		if (!MyIsNumeric(fld.value))
		{
			alert(txtValNumVal);
			fld.value = '';
			fld.focus();
		}
		else if (dType == 'int')
		{
			fld.value = parseInt(fld.value);
		}
	}
}
var SaveImgField;
var SaveImgImage;
var SaveImgMaxSize;
function getImg(Field, Img, MaxSize)
{
	SaveImgField = Field;
	SaveImgImage = Img;
	SaveImgMaxSize = MaxSize;
	Start('../upload/fileupload.aspx?ID=' + dbID + '&style=../design/0/style/stylePopUp.css',300,100,'no')
}

function Start(page, w, h, s) {
OpenWin = this.open(page, "ImagePicker", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable=no, width="+w+",height="+h);
}

function changepic(img_src) {
SaveImgField.value = img_src;
SaveImgImage.src = "../pic.aspx?filename=" + img_src + "&MaxSize=" + SaveImgMaxSize + '&dbName=' + dbName;
}

function updateNote() 
{
	document.form1.LineMemo.value = document.form1.NoteVar.options[document.form1.NoteVar.selectedIndex].value
}

var objField;
function datePicker(page, w, h, s, r, o) 
{
objField = o
OpenWin = this.open(page, "datePicker", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable="+r+", width="+w+",height="+h);
OpenWin.focus()
}
function setTimeStamp(Action, varDate) { 
objField.value = varDate }

function chkThis(Field, FType, EditType, FSize)
{
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
			switch (jQuery.trim(EditType))
			{
				case "":
					if (Field.value != '')
					{
						if (!MyIsNumeric(getNumeric(Field.value)))
						{
							Field.value = "";
							alert(txtValNumVal);
						}
						else if (parseFloat(Field.value) < 1)
						{
							Field.value = "";
							alert(txtValNumMinVal.replace("{0}", "1"));
						}
						else if (parseFloat(Field.value) > 2147483647)
						{
							alert(txtValNumMaxVal.replace("{0}", "2147483647"));
							Field.value = 2147483647;
						}
						else if (parseFloat(getNumeric(Field.value))-parseInt(getNumeric(Field.value)) != 0)
						{
							Field.value = "";
							alert(txtValNumValWhole);
						}
					}
				case "T":
					break;
			}
			break;
		case 'B':
			if (Field.value != '')
			{
				if (!MyIsNumeric(getNumeric(Field.value)))
				{
					Field.value = "";
					alert(txtValNumVal);
				}
				else
				{
					if (parseFloat(getNumeric(Field.value)) > 1000000000000)
					{
						Field.value = 999999999999;
					}
					else if (parseFloat(getNumeric(Field.value)) < -1000000000000)
					{
						Field.value = -999999999999;
					}
					switch (jQuery.trim(EditType))
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
			}
			break;
	}
}
