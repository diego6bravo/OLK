var endNum;

function openPrint(secID, linkData)
{
	OpenWin = window.open('sectionPDF.asp?secID=' + secID + '&' + linkData.replace('{0}', endNum), 'Print', 'toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes, width=760,height=540');
	OpenWin.focus();
	//OpenWin.print();
}

function checkSubmitStatus()
{
	$.post("submitControlFetch.asp?d=" + (new Date()).toString(), { Action: 0, LogNum: LogNum, LogNumID: LogNumID },
   function(data){
     setSubmitStatus(data);
   });
}
function setSubmitStatus(data)
{
	try
	{
		var result = data.split('{S}');
		var status = result[0];
		var errCode = result[1];
		var errMsg = result[2];
		var objCode = result[3];
		var poolCount = result[4];
		endNum = result[5];
		
		switch (status)
		{
			case 'C':
				document.getElementById('submitStatus').innerHTML = txtTransactinoPool.replace('{0}', '<span class="' + poolCSS + '">' + poolCount + '</span>');
				setTimeout('checkSubmitStatus();', 2000);
				break;
			case 'H':
			case 'S':
				if (LogNum2 == null)
				{
					document.getElementById('submitTitle').innerText = txtTransComp;
					document.getElementById('submitStatus').innerText = transOkMsg.replace('{0}', endNum);
					document.getElementById('trSubmitOk').style.display = goEndFunc != '' ? '' : 'none';
					if (goSecondDesc != '')
					{
						document.getElementById('btnOk2').onclick = new Function(goSecondFunc.replace('{0}', endNum.replace('\'', '\\\'')));
						document.getElementById('trSubmitOk2').style.display = '';
					}
					if (goThirdDesc != '')
					{
						document.getElementById('btnOk3').onclick = new Function(goThirdFunc.replace('{0}', endNum.replace('\'', '\\\'')));
						document.getElementById('trSubmitOk3').style.display = '';
					}
					if (hasPrint && strPrint)
					{
						var arrPrint = strPrint.split(', ');
						for (var i = 0;i<arrPrint.length;i++)
						{
							var trSubmitPrint = document.getElementById('trSubmitPrint' + arrPrint[i]);
							trSubmitPrint.style.display = '';	
						}
					}
					document.getElementById('submitImgWait').style.display = 'none';
					document.getElementById('submitImgOK').style.display = '';
					document.getElementById('trSubmitTransPool').style.display = '';
					if (enableBG) document.getElementById('trSubmitRunInBG').style.display = 'none';
				}
				else
				{
					LogNum = LogNum2;
					LogNumID = LogNumID2;
					LogNum2 = null;
					LogNumID2 = '';
					setTimeout('checkSubmitStatus();', 1000);
				}
				break;
			case 'E':
				document.getElementById('submitTitle').innerText = txtError;
				document.getElementById('submitStatus').innerText = errMsg;
				document.getElementById('trSubmitTransPool').style.display = '';
				document.getElementById('trSubmitOk').style.display = '';
				document.getElementById('trSubmitErr').style.display = '';
				document.getElementById('submitImgWait').style.display = 'none';
				document.getElementById('submitImgError').style.display = '';
				if (enableBG) document.getElementById('trSubmitRunInBG').style.display = 'none';
				break;
			case 'P':
				document.getElementById('submitTitle').innerText = txtProcessing;
				document.getElementById('trSubmitTransPool').style.display = 'none';
				setTimeout('checkSubmitStatus();', 1000);
				break;
			default:
				alert('Submit control does not recognize the object status');
				break;
		}
	}
	catch (err)
	{
		alert('Error checking status: ' + err.description);
	}
}
setTimeout('checkSubmitStatus();', 2000);
function retryAction()
{
	$.post("submitControlFetch.asp?d=" + (new Date()).toString(), { Action: 1, LogNum: LogNum },
   function(data){
     setRetryAction(data);
   });
}
function setRetryAction(data)
{
	if (data == 'OK')
	{
		
		document.getElementById('submitTitle').innerText = txtWait + '...';
		document.getElementById('trSubmitErr').style.display = 'none';
		document.getElementById('trSubmitOk').style.display = 'none';
		document.getElementById('submitImgWait').style.display = '';
		document.getElementById('submitImgError').style.display = 'none';
		if (enableBG) document.getElementById('trSubmitRunInBG').style.display = '';
		checkSubmitStatus();
	}
	else
	{
		alert('Error retry transtaction');
	}
}
function runInBG()
{
	$.post("submitControlFetch.asp?d=" + (new Date()).toString(), { Action: 2, LogNum: LogNum, LogNumID: LogNumID },
   function(data){
     setRunInBG(data);
   });
}

function setRunInBG(data)
{
	if (data == 'OK')
	{
		window.location.href = bgRedir;
	}
	else
	{
		alert('Error retry transtaction');
	}
}
