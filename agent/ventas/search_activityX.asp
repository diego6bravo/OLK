<% addLngPathStr = "ventas/" %>
<!--#include file="lang/search_activityX.asp" -->
<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>
<script language="javascript">
var retAcct
function Start(page, retAction) {
retAcct = retAction
OpenWin = this.open(page, "DatePicker", "toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no, width=240,height=220");
}

function setTimeStamp(retAction, varDate) 
{
	retAcct.value = varDate;
}

function getValue(myType, fld) 
{
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
<form method="POST" action="searchOpenedActivities.asp" name="frmSmallSearch">
<input type="hidden" name="cmd" value="searchOpenedActivities">
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
					<p align="left"><%=getsearch_activityXLngStr("LttlPendActSearch")%></td>
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
											<td width="79"><%=getsearch_activityXLngStr("DtxtSource")%></td>
											<td>
											<select name="cmbSourceType" size="1">
											<option value=""><%=getsearch_activityXLngStr("DtxtAll")%></option>
											<option value="O"><%=getsearch_activityXLngStr("DtxtOLK")%></option>
											<option value="S"><%=getsearch_activityXLngStr("DtxtSAP")%></option>
											</select></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td width="79"><%=getsearch_activityXLngStr("DtxtLogNum")%></td>
											<td>
											<input type="text" name="LogNumFrom" size="11" onkeydown="return valKeyNum(event);" onchange="if(document.frmSmallSearch.LogNumTo.value=='')document.frmSmallSearch.LogNumTo.value=this.value;" onfocus="this.select();"> 
											-
											<input type="text" name="LogNumTo" size="11" onfocus="this.select();" onkeydown="return valKeyNum(event);"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td width="79"><%=getsearch_activityXLngStr("DtxtDate")%></td>
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
										<% If myAut.HasAuthorization(97) Then %>
										<tr class="GeneralTblBold2">
											<td width="79"><% If 1 = 2 Then %>Agents<% Else %><%=txtAgent%><% End If %></td>
											<td>
											<input type="text" name="SlpCodeFrom" size="15" onkeydown="return chkMax(event, this, 32);" onchange="javascript:getValue('Slp', this);" onfocus="this.select();"> 
											-
											<input type="text" name="SlpCodeTo" size="15" onkeydown="return chkMax(event, this, 32);" onchange="javascript:getValue('Slp', this);" onfocus="this.select();"></td>
										</tr>
										<% Else %>
										<input type="hidden" name="SlpCodeFrom" value="<%=myHTMLEncode(AgentName)%>">
										<input type="hidden" name="SlpCodeTo" value="<%=myHTMLEncode(AgentName)%>">
										<% End If %>
										<tr class="GeneralTblBold2">
											<td width="79"><%=getsearch_activityXLngStr("LtxtAsignedTo")%></td>
											<td>
											<input type="text" name="AttendUserFrom" size="15" onkeydown="return chkMax(event, this, 32);" onchange="javascript:getValue('Usr', this);" onfocus="this.select();"> 
											-
											<input type="text" name="AttendUserTo" size="15" onkeydown="return chkMax(event, this, 32);" onchange="javascript:getValue('Usr', this);" onfocus="this.select();"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td width="79"><% If 1 = 2 Then %><%=getsearch_activityXLngStr("DtxtClient")%><% Else %><%=txtClient%><% End If %></td>
											<td>
											<input type="text" name="CardCodeFrom" size="15" onkeydown="return chkMax(event, this, 15);" onchange="javascript:getValue('Crd', this);" onfocus="this.select();"> 
											-
											<input type="text" name="CardCodeTo" size="15" onkeydown="return chkMax(event, this, 15);" onchange="javascript:getValue('Crd', this);" onfocus="this.select();"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td width="79"><%=getsearch_activityXLngStr("DtxtGroup")%></td>
											<td>
											<input type="text" name="GroupNameFrom" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('Grp', this);" onfocus="this.select();"> 
											-
											<input type="text" name="GroupNameTo" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('Grp', this);" onfocus="this.select();"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td width="79"><%=getsearch_activityXLngStr("DtxtCountry")%></td>
											<td>
											<input type="text" name="CountryFrom" size="20" onkeydown="return chkMax(event, this, 100);" onchange="javascript:getValue('Cty', this);" onfocus="this.select();"> 
											-
											<input type="text" name="CountryTo" size="20" onkeydown="return chkMax(event, this, 100);" onchange="javascript:getValue('Cty', this);" onfocus="this.select();"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td width="79"><%=getsearch_activityXLngStr("DtxtActivity")%></td>
											<td>
											<select size="1" name="DocType">
										<option></option>
										<option value="C"><%=getsearch_activityXLngStr("DtxtConv")%></option>
										<option value="M"><%=getsearch_activityXLngStr("DtxtMeeting")%></option>
										<option value="E"><%=getsearch_activityXLngStr("DtxtNote")%></option>
										<option value="O"><%=getsearch_activityXLngStr("DtxtOther")%></option>
										<option value="T"><%=getsearch_activityXLngStr("DtxtTask")%></option>
										</select></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td width="79"><%=getsearch_activityXLngStr("DtxtOrder")%></td>
											<td>
											<select size="1" name="orden1">
									<option value="1"><%=getsearch_activityXLngStr("DtxtLogNum")%></option>
									<option value="SlpName"><%=txtAgent%></option>
									<option value="C1.CardCode"><%=getsearch_activityXLngStr("DtxtCode")%></option>
									<option value="CardName"><%=getsearch_activityXLngStr("DtxtName")%></option>
									<option value="GroupName"><%=getsearch_activityXLngStr("DtxtGroup")%></option>
									<option value="Country"><%=getsearch_activityXLngStr("DtxtCountry")%></option>
									<option value="CntctDateSort"><%=getsearch_activityXLngStr("LtxtCntDate")%></option>
									<option value="Action"><%=getsearch_activityXLngStr("DtxtActivity")%></option>
									</select> <select size="1" name="orden2">
									<option value="asc"><%=getsearch_activityXLngStr("DtxtAsc")%></option>
									<option value="desc" selected><%=getsearch_activityXLngStr("DtxtDesc")%></option>
									</select></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td colspan="2">
											<p align="center">
											<input type="submit" value="<%=getsearch_activityXLngStr("DbtnSearch")%>" name="B1"></td>
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
		<tr>
			<td valign="top">
			<iframe id="ifGetValue" name="ifGetValue" style="display: none" height="145" width="160" src=""></iframe>
			&nbsp;</td>
		</tr>
		</table>
</div>
</form>
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