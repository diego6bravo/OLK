function valFrmIndex()
{
	NavIndexByX = document.frmNavIndex.NavIndexByX.value;
	NavIndexByY = document.frmNavIndex.NavIndexByY.value;

	if (NavIndexByX > 1 || NavIndexByY > 1)
	{
		myCmb = document.frmNavIndex.cmbNavID;
		for (var i = 0;i<myCmb.length;i++)
		{
			for (var j = 0;j<myCmb.length;j++)
			{
				if (j != i && myCmb[i].selectedIndex == myCmb[j].selectedIndex && myCmb[i].selectedIndex != 0)
				{
					alert(txtValEqIndex.replace('{0}', (i+1)).replace('{1}', (j+1)));
					return false;
				}
			}
		}
	}
	return true;
}

function doIndex()
{
	myCmb = document.frmNavIndex.cmbNavID
	var SelVals = '';
	for (var i = 0;i<myCmb.length;i++)
	{
		if (i > 0) SelVals += ',';
		SelVals += myCmb[i].selectedIndex;
	}
	
	NavIndexByX = document.frmNavIndex.NavIndexByX.value;
	NavIndexByY = document.frmNavIndex.NavIndexByY.value;
	
	strTbl = 	'<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" width=\"500\">';
	
	var x = 1;
	for (var i = 0;i<NavIndexByY;i++)
	{
		strTbl +=	'<tr>';
		
		for (var j = 0;j<NavIndexByX;j++)
		{
			strTbl += '<td align=\"center\" class=\"tdIndex\">';
			
			strTbl += x + '<br>' + getCmbNav(x);
			
			strTbl += '</td>';
			
			x++;
		}
		
		strTbl += 	'</tr>';
	}
	
	strTbl += 	'</table>';
	
	tdIndex.innerHTML = strTbl;
	
	
	myCmb = document.frmNavIndex.cmbNavID
	arrVals = SelVals.split(',');
	for (var i = 0;i<myCmb.length;i++)
	{
		if (i < arrVals.length) myCmb[i].selectedIndex = arrVals[i]
	}

}

function getCmbNav(x)
{
	strCmb = '<select name=\"NavID' + x + '\" id=\"cmbNavID\" size=\"1\" class=\"input\">';
	//strCmb += '<option></option>';
	
	if (myNav != '')
	{
		if (myNav.length)
		{
			for (var k = 0;k<myNav.length;k++)
			{
				strCmb += '<option value=\"' + myNav[k].split('|')[0] + '\">' + myNav[k].split('|')[1] + '</option>';
			}
		}
		else
		{
			strCmb += '<option value=\"' + myNav.split('|')[0] + '\">' + myNav.split('|')[1] + '</option>';
		}
	}
	
	strCmb += '</select>';
	
	return strCmb;
}