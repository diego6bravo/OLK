function valFrm()
{
	myCmb = document.frmSecIndex.cmbSecID;
	SecIndexByX = document.frmSecIndex.SecIndexByX;
	SecIndexByY = document.frmSecIndex.SecIndexByY;
	
	if (parseInt(SecIndexByX.value) > 1 && parseInt(SecIndexByY.value) > 1)
	{
		for (var i = 0;i<myCmb.length;i++)
		{
			if (myCmb[i].selectedIndex < 1)
			{
				alert(txtValSelIndex.replace('{0}', (i+1)));
				return false;
			}
			else
			{
				for (var j = 0;j<myCmb.length;j++)
				{
					if (j != i && myCmb[i].selectedIndex == myCmb[j].selectedIndex)
					{
						alert(txtValEqIndex.replace('{0}', (i+1)).replace('{1}', (j+1)));
						return false;
					}
				}
			}
		}
	}
	else
	{
		if (myCmb.selectedIndex < 1)
		{
			alert(txtValSelIndex.replace('{0}', 1));
			return false;
		}
	}
	return true;
}

function doIndex()
{
	myCmb = document.frmSecIndex.cmbSecID
	var SelVals = '';
	for (var i = 0;i<myCmb.length;i++)
	{
		if (i > 0) SelVals += ',';
		SelVals += myCmb[i].selectedIndex;
	}
	
	SecIndexByX = document.frmSecIndex.SecIndexByX.value;
	SecIndexByY = document.frmSecIndex.SecIndexByY.value;
	
	strTbl = 	'<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" width=\"500\">';
	
	var x = 1;
	for (var i = 0;i<SecIndexByY;i++)
	{
		strTbl +=	'<tr>';
		
		for (var j = 0;j<SecIndexByX;j++)
		{
			strTbl += '<td align=\"center\" class=\"tdIndex\">';
			
			strTbl += x + '<br>' + getCmbSec(x);
			
			strTbl += '</td>';
			
			x++;
		}
		
		strTbl += 	'</tr>';
	}
	
	strTbl += 	'</table>';
	
	tdIndex.innerHTML = strTbl;
	
	
	myCmb = document.frmSecIndex.cmbSecID
	arrVals = SelVals.split(',');
	for (var i = 0;i<myCmb.length;i++)
	{
		if (i < arrVals.length) myCmb[i].selectedIndex = arrVals[i]
	}

}

function getCmbSec(x)
{
	strCmb = '<select name=\"SecID' + x + '\" id=\"cmbSecID\" size=\"1\" class=\"cmbSec\">';
	strCmb += '<option></option>';
	
	if (mySec != '')
	{
		if (mySec.length)
		{
			for (var k = 0;k<mySec.length;k++)
			{
				strCmb += '<option value=\"' + mySec[k].split('|')[0] + '\">' + mySec[k].split('|')[1] + '</option>';
			}
		}
		else
		{
			strCmb += '<option value=\"' + mySec.split('|')[0] + '\">' + mySec.split('|')[1] + '</option>';
		}
	}
	
	strCmb += '</select>';
	
	return strCmb;
}