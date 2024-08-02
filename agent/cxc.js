function printStory() {

	var trCXCCmpName = document.getElementById('trCXCCmpName') && document.getElementById('hdTitle');
	
	w=window.open('','newwin')
	w.document.write('<html ' + (rtl != '' ? 'dir="rtl"' : '') + '><head><link rel="stylesheet" type="text/css" href="design/' + SelDes + '/style/stylenuevo.css"></head><body onLoad="window.print()">');
   
	if (trCXCCmpName)
	{
		w.document.write('<table cellpadding="0" cellspacing="0" border="0" width="100%">' + document.getElementById('hdTitle').value + '</table>');
	}
	
	w.document.write(document.getElementById('dvCXCPrint').innerHTML);
   
	w.document.write('</body></html>');

	controls = w.document.getElementsByTagName('input');
	for (var i = 0;i<controls.length;i++)
	{
		if (controls[i].attributes.item('type').value == 'button' || controls[i].attributes.item('type').value == 'submit')
		{
			controls[i].style.display = 'none';
		}
	}
	
	if (trCXCCmpName)
	{
		w.document.getElementById('trCXCCmpName').style.display = 'none';
	}
	
	w.document.close();
}
