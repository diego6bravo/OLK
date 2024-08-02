<% GrpLevel = CInt(Request("GrpLevel")) %>
function valShow()
{
	if (document.frmMain.LvlTyp1.selectedIndex == 0)
	{
		alert(txtValSelDimX.replace('{0}', '1'));
		return false;
	}
	<% If GrpLevel >= 2 Then %>
	else if (document.frmMain.LvlTyp2.selectedIndex == 0)
	{
		alert(txtValSelDimX.replace('{0}', '2'));
		return false;
	}
	else if (document.frmMain.LvlTyp1.selectedIndex == document.frmMain.LvlTyp2.selectedIndex)
	{
		alert(txtValRepDim);
		return false;
	}
	<% End If %>
	return true;
}

function chkQty(fld, SelQty, MaxQty)
{
	if (!IsNumeric(fld.value) || fld.value == '')
	{
		alert(txtValNumVal);
		fld.value = SelQty;
		fld.focus();
	}
	else if (parseFloat(fld.value) < 0)
	{
		alert(txtValNumMinVal.replace('{0}', '0'));
		fld.value = SelQty;
		fld.focus();
	}
	else if (parseFloat(fld.value) > parseFloat(MaxQty))
	{
		alert(txtValMoreThenAvl);
		fld.value = MaxQty;
		fld.focus();
	}
	SumTotal();
}

function SumTotal()
{
	selQty = document.frmMain.SelQty;
	var sQtyCount = 0;
	if (selQty.length)
	{
		for (var i = 0;i<selQty.length;i++)
		{
			sQtyCount += parseFloat(selQty(i).value);
		}
	}
	else
	{
		sQtyCount = parseFloat(selQty.value);
	}
	document.frmMain.txtSelQty.value = sQtyCount;
	document.frmMain.txtOpenQty.value = parseFloat(document.frmMain.txtReqQty.value)-sQtyCount ;
}

function IsNumeric(sText)
{
   var ValidChars = "0123456789.-";
   var IsNumber=true;
   var Char;

 
   for (i = 0; i < sText.length && IsNumber == true; i++) 
      { 
      Char = sText.charAt(i); 
      if (ValidChars.indexOf(Char) == -1) 
         {
         IsNumber = false;
         }
      }
   return IsNumber;
   
}

function Right(str, n){
    if (n <= 0)
       return "";
    else if (n > String(str).length)
       return str;
    else {
       var iLen = String(str).length;
       return String(str).substring(iLen, iLen - n);
    }
}

function valFrm()
{
	if (parseFloat(document.frmMain.txtOpenQty.value) < 0)
	{
		alert(txtValSelOverQty);
		return false;
	}
	return true;
}

var OpenWin = this
function Start(page, w, h, s) {
OpenWin = this.open(page, "ImageThumb", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable=yes, width="+w+",height="+h);
}

function chOtherDim()
{
	var chkVal = '';
	for (var i = 0;i<=5;i++)
	{
		if (i > 0) chkVal += '{|}';
		if (document.getElementById('LvlSelVal' + i) != null) 
			chkVal += document.getElementById('LvlSelVal' + i).value;
	}
	
	var arrChkVal = chkVal.split('{|}');
	var sumRow = new Array(tblMatrix.rows.length-2);
	for (var i = 0;i<sumRow.length;i++) sumRow[i] = 0;
	var sumCol = new Array(colCount.length);
	for (var i = 0;i<sumCol.length;i++) sumCol[i] = 0;
	for (var i = 0;i<tblMatrix.rows.length;i++)
	{
		colI = 0;
		if (i > 1)
		for (var j = 1;j<tblMatrix.rows(i).cells.length;j++)
		{
			c = tblMatrix.rows(i).cells(j);
			if (c.id.length > 0)
			{
				arrChk = Right(c.id, c.id.length).split('-');
				hide = false;
				for (var v = 0;v<arrChkVal.length;v++)
				{
					if (arrChkVal[v] != '' && arrChkVal[v] != arrChk[v])
					{
							hide = true;
							break;
					}
				}
				if (hide) c.style.display = 'none';
				else 
				{
					c.style.display = '';
					if (IsNumeric(c.innerHTML))
					{
						sumRow[i-2] += parseFloat(c.innerHTML);
						sumCol[colI++] += parseFloat(c.innerHTML);
					}
				}
			}
		}
	}

	for (var i = 0;i<sumRow.length;i++)
	{
		if (sumRow[i] <= 0)
			tblMatrix.rows(i+2).style.display = 'none';
		else
			tblMatrix.rows(i+2).style.display = '';
	}

	if (document.frmMain.LvlTyp2)
	{
		selTyp2 = parseInt(document.frmMain.LvlTyp2.value);
		
		for (var i = 0;i<sumCol.length;i++)
		{
			if (sumCol[i] <= 0)
			{
				tblMatrix.rows(0).cells.item(i+1).style.display = 'none';
				tblMatrix.rows(1).cells.item(i*2).style.display = 'none';
				tblMatrix.rows(1).cells.item((i*2)+1).style.display = 'none';

				for (var j = 2;j<tblMatrix.rows.length;j++)
				{
					for (var k = 2;k<tblMatrix.rows(j).cells.length;k++)
					{
						if (tblMatrix.rows(j).cells(k).id.split('-')[selTyp2] == colCount[i])
						{
							tblMatrix.rows(j).cells(k).style.display = 'none';
						}
					}
				}
			}
			else
			{
				tblMatrix.rows(0).cells.item(i+1).style.display = '';
				tblMatrix.rows(1).cells.item(i*2).style.display = '';
				tblMatrix.rows(1).cells.item((i*2)+1).style.display = '';
			}
		}
	}
}
