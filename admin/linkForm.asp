<form name="frmMyLink" method="post" action="">
</form>
<script language="javascript">
function doMyLink(action, vars, target)
{
	document.frmMyLink.innerHTML = "";
	document.frmMyLink.action = action;
	document.frmMyLink.target = target;
	var arrVars = vars.split('&');
	for (var v = 0;v<arrVars.length;v++)
	{
		document.frmMyLink.innerHTML += '<input type=\"hidden\" name=\"' + arrVars[v].split('=')[0] + '\" value=\"' + arrVars[v].split('=')[1] + '\">';
	}
	document.frmMyLink.submit();
}
</script>