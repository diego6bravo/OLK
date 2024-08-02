function forceRefreshMon()
{
	window.clearTimeout(taskMonTimerID);
	refreshMon()
}

function refreshMon()
{
	strGetID = "";

	for (var i = 0;i<TaskMonType.length;i++)
	{
		if (strGetID != "") strGetID = strGetID + ","
		strGetID = strGetID + TaskMonType[i].value + '|' + TaskMonID[i].value;
	}

	var url='taskMonitorFetch.asp?GetID=' + strGetID;

	xmlHttp=GetXmlHttpObject(setTaskMonValues);
	xmlHttp.open("GET", url , true);
	xmlHttp.send(null);
}

function setTaskMonValues()
{
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
	{
		var arrValues = xmlHttp.responseText.split('{S}');
		
		for (var i = 0;i<arrValues.length;i++)
		{
			var values = arrValues[i].split('|');
			
			var vType = values[0];
			var id = parseInt(values[1]);
			var value = values[2];
			
			document.getElementById('tdTaskMonVal' + vType + id).innerHTML = value;

			switch (vType)
			{
				case 'S':
					switch (id)
					{
						case 3:
							if (value != -1)
							{
								var source = values[3];
								var card = values[4];
								
								switch (source)
								{
									case 'O':
										document.frmGoAct.action = 'addActivity/goActivity.asp';
										document.frmGoAct.LogNum.value = value;
										document.frmGoAct.Card.value = card;
										
										break;
									case 'S':
										document.frmGoAct.action = 'addActivity/goEditActivity.asp';
										document.frmGoAct.ClgCode.value = value;
										document.frmGoAct.CardCode.value = card;
										break;
								}
							}
							document.getElementById('imgTaskMonS3').src = 'images/icon_activity_' + source + '.gif';
							document.getElementById('trTaskMonS3').style.display = (value == -1) ? 'none' : '';
							break;
						case 0:
						case 7:
						case 8:
						case 9:
						case 10:
						case 11:
						case 12:
						case 13:
						case 14:
						case 15:
						case 16:
						case 17:
							document.getElementById('trTaskMonS' + id).style.display = (value == 0) ? 'none' : '';
							break;
					}
					break;
				case 'U':
					var doHide = document.getElementById('TaskMonHideNullU' + id).value == 'Y';
					if (doHide)
					{
						document.getElementById('trTaskMonU' + id).style.display = (value == '' || value == '0') ? 'none' : '';
					}
					break;
						
			}
		}
		
		taskMonTimerID = window.setTimeout('forceRefreshMon();', 60000);
	}
}

var taskMonTimerID = window.setTimeout('forceRefreshMon();', 60000);

