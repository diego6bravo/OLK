var maxOrdr = 0;
var lastSlp = -1;
function delGrp(id)
{
	if(confirm(txtConfDelGrp.replace('{0}', document.getElementById('GrpName' + id).value)))
		window.location.href='adminSubmit.asp?submitCmd=AutGrp&cmd=delGrp&delIndex=' + id;
}

function doAddSlp(slp)
{
	var slpCode = slp.value;
	var slpName = slp.options[slp.selectedIndex].innerText;
	
	var str = '';
	str += '					<tr>' ;
	str += '						<td class="TblRepNrm"><input type="hidden" name="SlpCode" value="' + slpCode + '">' + slpName + '</td>' ;
	str += '						<td class="TblRepNrm"><select name="Op' + slpCode + '" id="Op' + slpCode + '" style="display: none;">' ;
	str += '						<option value="O" selected>' + txtOr + '</option>' ;
	str += '						<option value="A">' + txtAnd + '</option>' ;
	str += '						</select>' ;
	str += '						</td>' ;
	str += '						<td align="center"><table cellpadding="0" border="0">' ;
	str += '								<tr>' ;
	str += '									<td class="TblRepNrm"><input type="text" name="Order' + slpCode + '" id="Order' + slpCode + '" size="5" style="text-align:right" class="input" value="' + (maxOrdr++) + '" onfocus="this.select()" onchange="chkThis(this, 0, 0)" onkeydown="return chkMax(event, this, 6);"></td>' ;
	str += '									<td valign="middle">' ;
	str += '									<table cellpadding="0" cellspacing="0" border="0">' ;
	str += '										<tr>' ;
	str += '											<td><img src="images/img_nud_up.gif" id="btnOrder' + slpCode + 'Up"></td>' ;
	str += '										</tr>' ;
	str += '										<tr>' ;
	str += '											<td><img src="images/spacer.gif"></td>' ;
	str += '										</tr>' ;
	str += '										<tr>' ;
	str += '											<td><img src="images/img_nud_down.gif" id="btnOrder' + slpCode + 'Down"></td>' ;
	str += '										</tr>' ;
	str += '									</table>' ;
	str += '									</td>' ;
	str += '								</tr>' ;
	str += '							</table>' ;
	str += '						</td><td class="TblRepNrm"><img border="0" src="images/remove.gif" width="16" height="16" onclick="delSlp(this, ' + slpCode + ');"></td>' ;
	str += '					</tr>' ;

	$("#tblSlp tr:last").before(str);

	NumUDAttach('frmEditGroups', 'Order' + slpCode, 'btnOrder' + slpCode + 'Up', 'btnOrder' + slpCode + 'Down');
	
	if (lastSlp != -1 && document.getElementById('Op' + lastSlp)) document.getElementById('Op' + lastSlp).style.display = '';
	
	lastSlp = slpCode;
	
	slp.options.remove(slp.selectedIndex);
	slp.selectedIndex = 0;
	
}

function delSlp(slp, delSlpCode)
{
	if (confirm(txtConfDel))
	{
		$(slp).parent().parent().remove();
		
		delID = document.frmEditGroups.delID;
		if (delID.value != '') delID.value += ', ';
		delID.value += delSlpCode;

	}
}