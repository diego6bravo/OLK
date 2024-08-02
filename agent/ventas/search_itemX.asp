<% addLngPathStr = "ventas/" %>
<!--#include file="lang/search_itemX.asp" -->


<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>
<script language="javascript">
var retAcct
function Start(page, retAction) {
retAcct = retAction
OpenWin = this.open(page, "DatePicker", "toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no, width=240,height=220");
}
// End -->
function setTimeStamp(retAction, varDate) {
retAcct.value = varDate }
</script>

<script language="javascript">
function getValue(myType, fld) {
if (fld.value == '') { return; } 
	updFld = fld;
	if (fld.value.indexOf('*') == -1) {
		document.frmGetValue.Type.value = myType;
		document.frmGetValue.searchStr.value = fld.value;
		document.frmGetValue.submit();
	}
	else { launchSelect(myType, fld.value); }
}
function launchSelect(myType, Value){
	var retVal = window.showModalDialog('topGetValueSelect.asp?Type=' + myType + '&Value=' + Value,'','dialogWidth:500px;dialogHeight:500px');
	if (retVal != '' && retVal != null){
		updFld.value = retVal; setTargetVal(retVal); retVal = '';
	} 
	else { 
		updFld.value = '';
	}
}
function setValue(src, value, myType){
	if (value != '') 
	{ updFld.value = value; setTargetVal(value); }
	else { if(src == 0)launchSelect(myType, updFld.value); }
}
function setTargetVal(value)
{
	if (Right(updFld.name, 4) == "From")
	{
		setFldName = Left(updFld.name, (updFld.name.length-4));
		fldTo = document.getElementById(setFldName + 'To');
		if (fldTo.value == '') { fldTo.value = value; fldTo.select(); }
	}
}
</script>
<form method="post" target="ifGetValue" name="frmGetValue" action="topGetValue.asp">
<input type="hidden" name="Type" value="">
<input type="hidden" name="searchStr" value="">
</form>
<form method="POST" action="searchOpenedItems.asp" name="frmSmallSearch">
<input type="hidden" name="cmd" value="searchOpenedItems">
<div align="center">
	<table border="0" cellpadding="0" width="499" id="table1">
		<tr>
			<td>
			<p align="center">
			<img border="0" src="design/0/images/search_top.jpg" width="407" height="140"></td>
		</tr>
		<tr>
			<td valign="top">
			<table border="0" cellpadding="0" width="100%" id="table2">
				<tr class="GeneralTlt">
					<td>
					<p align="left"><%=getsearch_itemXLngStr("LttlPendITmSearch")%></td>
				</tr>
				<tr>
					<td>
					<table border="0" cellpadding="0" width="100%" cellspacing="0" id="table4">
						<tr class="GeneralTlt">
							<td>
							<table border="0" cellpadding="0" width="100%" cellspacing="1" id="table5">
								<tr class="GeneralTbl">
									<td>
									<table border="0" cellpadding="0" width="100%" cellspacing="1" id="table6">
										<tr class="GeneralTblBold2">
											<td width="79"><%=getsearch_itemXLngStr("DtxtLogNum")%></td>
											<td>
											<input type="text" name="LogNumFrom" size="11" onkeydown="return valKeyNum(event);" onchange="if(document.frmSmallSearch.LogNumTo.value=='')document.frmSmallSearch.LogNumTo.value=this.value;" onfocus="this.select();"> 
											-
											<input type="text" name="LogNumTo" size="11" onfocus="this.select();" onkeydown="return valKeyNum(event);"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td width="79"><%=getsearch_itemXLngStr("DtxtDate")%></td>
											<td>
											<table cellpadding="0" cellspacing="0" border="0">
												<tr>
													<td><input readonly name="dtFrom" size="11" onclick="btnDtFrom.click()"></td>
													<td>&nbsp;<img border="0" src="images/cal.gif" id="btnDtFrom"></td>
													<td>&nbsp;-&nbsp;</td>
													<td><input readonly name="dtTo" size="11" onclick="btnDtTo.click()"></td>
													<td>&nbsp;<img border="0" src="images/cal.gif" id="btnDtTo"></td>
												</tr>
											</table>
											</td>
										</tr>
										<tr class="GeneralTblBold2">
											<td width="79"><%=getsearch_itemXLngStr("DtxtItem")%></td>
											<td>
											<input type="text" name="ItemCodeFrom" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('TItm', this);" onfocus="this.select();"> 
											-
											<input type="text" name="ItemCodeTo" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('TItm', this);" onfocus="this.select();"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td width="79"><%=txtAlterGrp%></td>
											<td>
											<input type="text" name="ItmsGrpNamFrom" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('ItmGrp', this);" onfocus="this.select();"> 
											-
											<input type="text" name="ItmsGrpNamTo" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('ItmGrp', this);" onfocus="this.select();"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td width="79"><%=txtAlterFrm%></td>
											<td>
											<input type="text" name="FirmNameFrom" size="20" onkeydown="return chkMax(event, this, 30);" onchange="javascript:getValue('ItmFrm', this);" onfocus="this.select();"> 
											-
											<input type="text" name="FirmNameTo" size="20" onkeydown="return chkMax(event, this, 30);" onchange="javascript:getValue('ItmFrm', this);" onfocus="this.select();"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td width="79"><%=getsearch_itemXLngStr("DtxtOrder")%></td>
											<td>
											<select size="1" name="orden1">
										<option value="T0.LogNum"><%=getsearch_itemXLngStr("DtxtLogNum")%></option>
										<option value="T0.ItemCode"><%=getsearch_itemXLngStr("DtxtCode")%></option>
										<option value="T0.ItemName"><%=getsearch_itemXLngStr("DtxtDescription")%></option>
										<option value="T2.ItmsGrpNam"><%=txtAlterGrp%></option>
										<option value="T3.FirmName"><%=txtAlterFrm%></option>
										<option value="CreateDateSort"><%=getsearch_itemXLngStr("DtxtDate")%></option>
										<option value="StatusStr"><%=getsearch_itemXLngStr("DtxtState")%></option>
										</select> <select size="1" name="orden2">
										<option value="asc"><%=getsearch_itemXLngStr("DtxtAsc")%></option>
										<option value="desc" selected>
										<%=getsearch_itemXLngStr("DtxtDesc")%></option>
										</select></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td colspan="2">
											<p align="center">
											<input type="submit" value="<%=getsearch_itemXLngStr("DbtnSearch")%>" name="btnSearch"></td>
										</tr>
										</table>
									</td>
								</tr>
							</table>
							</td>
						</tr>
					</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
	</table>
</div>
</form>

<p align="center">
			<iframe id="ifGetValue" name="ifGetValue" style="display: none" height="99" width="256" src=""></iframe>
</p>
<script type="text/javascript">
    Calendar.setup({
        inputField     :    "dtFrom",     // id of the input field
        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
        button         :    "btnDtFrom",  // trigger for the calendar (button ID)
        align          :    "Bl",           // alignment (defaults to "Bl")
        singleClick    :    true
    });
    Calendar.setup({
        inputField     :    "dtTo",     // id of the input field
        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
        button         :    "btnDtTo",  // trigger for the calendar (button ID)
        align          :    "Br",           // alignment (defaults to "Bl")
        singleClick    :    true
    });
</script>
