<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Session("useraccess") <> "P" Then Response.Redirect "unauthorized.asp" %>
<% addLngPathStr = "" %>
<!--#include file="lang/recoverSearch.asp" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	direction: ltr;
}
</style>
</head>

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
<form method="POST" action="searchRecover.asp" name="frmSmallSearch">
<input type="hidden" name="cmd" value="searchRecoverDocs">
<div align="center">
	<table border="0" cellpadding="0" width="550" id="table1">
		<tr>
			<td colspan="2">
			<p align="center">
			<img border="0" src="design/0/images/search_top.jpg" width="407" height="140"></td>
		</tr>
		<tr>
			<td width="190" valign="top">&nbsp;</td>
			<td valign="top">
			<table border="0" cellpadding="0" width="100%">
				<tr class="GeneralTlt">
					<td>
					<p align="left"><%=getrecoverSearchLngStr("LttlRecDocSearch")%></td>
				</tr>
				<tr>
					<td>
					<table border="0" cellpadding="0" width="100%" cellspacing="0" >
						<tr class="GeneralTlt">
							<td>
							<table border="0" cellpadding="0" width="100%" cellspacing="1">
								<tr class="GeneralTbl">
									<td>
									<table border="0" cellpadding="0" width="100%" cellspacing="1">
										<tr class="GeneralTblBold2">
											<td><%=getrecoverSearchLngStr("DtxtLogNum")%></td>
											<td>
											<input type="text" name="LogNumFrom" size="11" onkeydown="return valKeyNum(event);" onchange="if(document.frmSmallSearch.LogNumTo.value=='')document.frmSmallSearch.LogNumTo.value=this.value;" onfocus="this.select();"> 
											-
											<input type="text" name="LogNumTo" size="11" onkeydown="return valKeyNum(event);" onfocus="this.select();"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td><%=getrecoverSearchLngStr("LtxtDate")%></td>
											<td>
											<table cellpadding="0" cellspacing="0" border="0">
												<tr>
													<td><input readonly name="dtFrom" size="11" onclick="btnDtFrom.click()"></td>
													<td><img border="0" src="images/cal.gif" id="btnDtFrom"></td>
													<td>-</td>
													<td><input readonly name="dtTo" size="11" onclick="btnDtTo.click()"></td>
													<td><img border="0" src="images/cal.gif" id="btnDtTo"></td>
												</tr>
											</table>
											</td>
										</tr>
										<tr class="GeneralTblBold2">
											<td><% If 1 = 2 Then %><%=getrecoverSearchLngStr("DtxtClients")%><% Else %><%=txtClients%><% End If %></td>
											<td>
											<input type="text" name="CardCodeFrom" size="13" onkeydown="return chkMax(event, this, 15);" onchange="javascript:getValue('Crd', this);" onfocus="this.select();"> 
											-
											<input type="text" name="CardCodeTo" size="13" onkeydown="return chkMax(event, this, 15);" onchange="javascript:getValue('Crd', this);" onfocus="this.select();"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td><%=getrecoverSearchLngStr("DtxtGroup")%></td>
											<td>
											<input type="text" name="GroupNameFrom" size="13" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('Grp', this);" onfocus="this.select();"> 
											-
											<input type="text" name="GroupNameTo" size="13" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('Grp', this);" onfocus="this.select();"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td><%=getrecoverSearchLngStr("DtxtGroup")%></td>
											<td>
											<input type="text" name="CountryFrom" size="13" onkeydown="return chkMax(event, this, 100);" onchange="javascript:getValue('Cty', this);" onfocus="this.select();"> 
											-
											<input type="text" name="CountryTo" size="13" onkeydown="return chkMax(event, this, 100);" onchange="javascript:getValue('Cty', this);" onfocus="this.select();"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td><%=getrecoverSearchLngStr("DtxtType")%></td>
											<td>
											<select size="1" name="DocType">
									<option></option>
									<option value="23"><% If 1 = 2 Then %>Cotizaciones<% Else %><%=txtQuotes%><% End If %></option>
									<option value="13"><% If 1 = 2 Then %>Facturas<% Else %><%=txtInvs%><% End If %></option>
									<option value="13C"><% If 1 = 2 Then %>Facturas<% Else %><%=txtInvs%><% End If %>/<% If 1 = 2 Then %>Recibo<% Else %><%=txtRct%><% End If %></option>
									<option value="24"><% If 1 = 2 Then %>Recibos<% Else %><%=txtRcts%><% End If %></option>
									<option value="17"><% If 1 = 2 Then %>Pedidos<% Else %><%=txtOrdrs%><% End If %></option>
									</select></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td class="style1">&nbsp;</td>
											<td>
											<input type="checkbox" style="background: background-image; border: 0px solid" name="SearchDel" id="SearchDel" value="ON"><label for="SearchDel"><%=getrecoverSearchLngStr("LtxtSearchDel")%></label></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td class="style1"><%=getrecoverSearchLngStr("DtxtOrder")%></td>
											<td>
											<select size="1" name="orden1">
											<option value="1">#<%=getrecoverSearchLngStr("DtxtLogNum")%></option>
											<option value="2"><%=getrecoverSearchLngStr("DtxtCode")%></option>
											<option value="7"><%=getrecoverSearchLngStr("DtxtDate")%></option>
											<option value="5"><%=getrecoverSearchLngStr("DtxtGroup")%></option>
											<option value="3"><%=getrecoverSearchLngStr("DtxtName")%></option>
											<option value="4"><%=getrecoverSearchLngStr("DtxtCountry")%></option>
											<option value="8"><%=getrecoverSearchLngStr("DtxtType")%></option>
											</select> <select size="1" name="orden2">
											<option value="asc"><%=getrecoverSearchLngStr("DtxtAsc")%></option>
											<option value="desc" selected>
											<%=getrecoverSearchLngStr("DtxtDesc")%></option>
											</select></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td colspan="2">
											<p align="center">
											<input type="submit" value="<%=getrecoverSearchLngStr("DtxtSearch")%>" name="B1"></td>
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

<iframe id="ifGetValue" name="ifGetValue" style="display: none" height="145" width="160" src=""></iframe>
<!--#include file="agentBottom.asp"-->