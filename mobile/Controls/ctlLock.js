function ClickUnlock(id)
{
	var imgLock = document.getElementById('imgLock' + id);
	var hdLockVal = document.getElementById('hdLockVal' + id);
	
	if (hdLockVal.value == 'Y')
	{
		hdLockVal.value = 'N';
		imgLock.src = imgLock.src.replace('lock', 'unlock');
	}
	else
	{
		hdLockVal.value = 'Y';
		imgLock.src = imgLock.src.replace('unlock', 'lock');
	}
}
