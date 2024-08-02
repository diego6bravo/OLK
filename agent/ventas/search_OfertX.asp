<% addLngPathStr = "ventas/" %>
<!--#include file="lang/search_OfertX.asp" -->
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
<form method="POST" action="ofertsMan.asp" name="frmSmallSearch">
<input type="hidden" name="cmd" value="ofertsMan">
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
					<p><% If 1 = 2 Then %><%=getsearch_OfertXLngStr("LttlSearchOferts")%><% Else %><%=Replace(getsearch_OfertXLngStr("LttlSearchOferts"), "{0}", txtOferts)%><% End If %></td>
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
											<td><%=getsearch_OfertXLngStr("DtxtDate")%></td>
											<td>
											<select size="1" name="dtBy">
											<option value="O"><% If 1 = 2 Then %>oferta<% Else %><%=txtOfert%><% End If %></option>
											<option value="R"><%=getsearch_OfertXLngStr("LtxtResponse")%></option>
											</select></td>
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
											<td><% If 1 = 2 Then %><%=getsearch_OfertXLngStr("DtxtClient")%><% Else %><%=txtClient%><% End If %></td>
											<td>
											&nbsp;</td>
											<td>
											<input type="text" name="CardCodeFrom" size="15" onkeydown="return chkMax(event, this, 15);" onchange="javascript:getValue('Crd', this);" onfocus="this.select();"> 
											-
											<input type="text" name="CardCodeTo" size="15" onkeydown="return chkMax(event, this, 15);" onchange="javascript:getValue('Crd', this);" onfocus="this.select();"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td><%=getsearch_OfertXLngStr("DtxtItem")%></td>
											<td>
											&nbsp;</td>
											<td>
											<input type="text" name="ItemCodeFrom" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('Itm', this);" onfocus="this.select();"> 
											-
											<input type="text" name="ItemCodeTo" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('Itm', this);" onfocus="this.select();"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td><%=getsearch_OfertXLngStr("DtxtPrice")%></td>
											<td>
											<select size="1" name="PriceBy">
											<option value="O"><% If 1 = 2 Then %>oferta<% Else %><%=txtOfert%><% End If %></option>
											<option value="R"><%=getsearch_OfertXLngStr("LtxtResponse")%></option>
											</select></td>
											<td>
											<input type="text" name="PriceFrom" size="13" onfocus="this.select();" onkeydown="return valKeyNumDec(event);" onchange="javascript:if(document.frmSmallSearch.PriceTo.value=='')document.frmSmallSearch.PriceTo.value=this.value;"> 
											-
											<input type="text" name="PriceTo" size="13" onfocus="this.select();" onkeydown="return valKeyNumDec(event);"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td><%=getsearch_OfertXLngStr("DtxtQty")%></td>
											<td>
											<select size="1" name="QtyBy">
											<option value="O"><% If 1 = 2 Then %>oferta<% Else %><%=txtOfert%><% End If %></option>
											<option value="R"><%=getsearch_OfertXLngStr("LtxtResponse")%></option>
											</select></td>
											<td>
											<input type="text" name="QtyFrom" size="13" onfocus="this.select();" onkeydown="return valKeyNumDec(event);" onchange="javascript:if(document.frmSmallSearch.QtyTo.value=='')document.frmSmallSearch.QtyTo.value=this.value;"> 
											-
											<input type="text" name="QtyTo" size="13" onfocus="this.select();" onkeydown="return valKeyNumDec(event);"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td><% If 1 = 2 Then %>Agents<% Else %><%=txtAgent%><% End If %></td>
											<td>
											&nbsp;</td>
											<td>
											<input type="text" name="SlpCodeFrom" size="20" onkeydown="return chkMax(event, this, 32);" onchange="javascript:getValue('Slp', this);" onfocus="this.select();"> 
											-
											<input type="text" name="SlpCodeTo" size="20" onkeydown="return chkMax(event, this, 32);" onchange="javascript:getValue('Slp', this);" onfocus="this.select();"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td><% If 1 = 2 Then %>Grupo de 
											Clientes<% Else %><%=Replace(getsearch_OfertXLngStr("LtxtClientGrps"), "{0}", txtClients)%><% End If %></td>
											<td>
											&nbsp;</td>
											<td>
											<input type="text" name="GroupNameFrom" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('Grp', this);" onfocus="this.select();"> 
											-
											<input type="text" name="GroupNameTo" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('Grp', this);" onfocus="this.select();"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td><%=getsearch_OfertXLngStr("DtxtCountry")%></td>
											<td>
											&nbsp;</td>
											<td>
											<input type="text" name="CountryFrom" size="20" onkeydown="return chkMax(event, this, 100);" onchange="javascript:getValue('Cty', this);" onfocus="this.select();"> 
											-
											<input type="text" name="CountryTo" size="20" onkeydown="return chkMax(event, this, 100);" onchange="javascript:getValue('Cty', this);" onfocus="this.select();"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td><%=txtAlterGrp%></td>
											<td>
											&nbsp;</td>
											<td>
											<input type="text" name="ItmsGrpNamFrom" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('ItmGrp', this);" onfocus="this.select();"> 
											-
											<input type="text" name="ItmsGrpNamTo" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('ItmGrp', this);" onfocus="this.select();"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td><%=txtAlterFrm%></td>
											<td>
											&nbsp;</td>
											<td>
											<input type="text" name="FirmNameFrom" size="20" onkeydown="return chkMax(event, this, 30);" onchange="javascript:getValue('ItmFrm', this);" onfocus="this.select();"> 
											-
											<input type="text" name="FirmNameTo" size="20" onkeydown="return chkMax(event, this, 30);" onchange="javascript:getValue('ItmFrm', this);" onfocus="this.select();"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td><%=getsearch_OfertXLngStr("DtxtNote")%></td>
											<td>
											<select size="1" name="NoteBy">
											<option value="O"><% If 1 = 2 Then %>oferta<% Else %><%=txtOfert%><% End If %></option>
											<option value="R"><%=getsearch_OfertXLngStr("LtxtResponse")%></option>
											</select></td>
											<td>
											<input type="text" name="Note" size="29" onkeydown="return chkMax(event, this, 100);" onfocus="this.select();"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td><%=getsearch_OfertXLngStr("LtxtDueDays")%></td>
											<td>
											<select size="1" name="DueBy">
											<option value="O"><% If 1 = 2 Then %>oferta<% Else %><%=txtOfert%><% End If %></option>
											<option value="R"><%=getsearch_OfertXLngStr("LtxtResponse")%></option>
											</select></td>
											<td>
											<input type="text" name="DueFrom" size="3" onchange="javascript:if(document.frmSmallSearch.DueTo.value=='')document.frmSmallSearch.DueTo.value=this.value;" onfocus="this.select();"> 
											-
											<input type="text" name="DueTo" size="3" onfocus="this.select();"></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td><%=getsearch_OfertXLngStr("DtxtState2")%></td>
											<td colspan="2">
											<table border="0" cellpadding="0" width="100%" id="table7">
												<tr class="GeneralTblBold2">
													<td>
													<input type="checkbox" name="OfertStatus" value="W" id="OfertStatusW" style="border-style:solid; border-width:0; background:background-image" checked><label for="OfertStatusW"><%=getsearch_OfertXLngStr("DtxtWaiting")%></label><br>
													<input type="checkbox" name="OfertStatus" value="A" id="OfertStatusA" style="border-style:solid; border-width:0; background:background-image"><label for="OfertStatusA"><%=getsearch_OfertXLngStr("DtxtAproved")%></label><br>
													<input type="checkbox" name="OfertStatus" value="R" id="OfertStatusR" style="border-style:solid; border-width:0; background:background-image"><label for="OfertStatusR"><%=getsearch_OfertXLngStr("DtxtRejected")%></label></td>
													<td>
													<input type="checkbox" name="OfertStatus" value="O" id="OfertStatusO" style="border-style:solid; border-width:0; background:background-image" checked><label for="OfertStatusO"><%=getsearch_OfertXLngStr("DtxtCounter")%>&nbsp;<% If 1 = 2 Then %>oferta<% Else %><%=txtOfert%><% End If %></label><br>
													<input type="checkbox" name="OfertStatus" value="B" id="OfertStatusB" style="border-style:solid; border-width:0; background:background-image"><label for="OfertStatusB"><%=getsearch_OfertXLngStr("LtxtPurchased")%></label><br>
													<input type="checkbox" name="OfertStatus" value="C" id="OfertStatusC" style="border-style:solid; border-width:0; background:background-image"><label for="OfertStatusC"><%=getsearch_OfertXLngStr("DtxtAnuled")%></label></td>
												</tr>
											</table>
											</td>
										</tr>
										<tr class="GeneralTblBold2">
											<td><%=getsearch_OfertXLngStr("DtxtOrder")%></td>
											<td colspan="2">
											<select size="1" name="orden1">
											<option value="8"><% If 1 = 2 Then %>Fecha 
											oferta<% Else %><%=Replace(getsearch_OfertXLngStr("LtxtOfertDate"), "{0}", txtOfert)%><% End If %></option>
											<option value="13"><%=getsearch_OfertXLngStr("LtxtRespDate")%></option>
											<option value="9"><% If 1 = 2 Then %>Fecha 
											Venc. oferta<% Else %><%=Replace(getsearch_OfertXLngStr("LtxtOfertDueDate"), "{0}", txtOfert)%><% End If %></option>
											<option value="14"><%=getsearch_OfertXLngStr("LtxtResDueDate")%></option>
											<option value="4"><%=getsearch_OfertXLngStr("DtxtItem")%></option>
											<option value="10"><% If 1 = 2 Then %>Cantidad 
											oferta<% Else %><%=Replace(getsearch_OfertXLngStr("LtxtOfertQty"), "{0}", txtOfert)%><% End If %></option>
											<option value="15"><%=getsearch_OfertXLngStr("LtxtRespQty")%></option>
											<option value="3"><% If 1 = 2 Then %><%=getsearch_OfertXLngStr("DtxtClient")%><% Else %><%=txtClients%><% End If %></option>
											<option value="7"><%=getsearch_OfertXLngStr("DtxtState")%></option>
											<option value="20"><% If 1 = 2 Then %>Ganancia 
											oferta<% Else %><%=Replace(getsearch_OfertXLngStr("LtxtOfertProfit"), "{0}", txtOfert)%><% End If %></option>
											<option value="21"><%=getsearch_OfertXLngStr("LtxtRespProfit")%></option>
											<option value="6"><%=getsearch_OfertXLngStr("LtxtBasePrice")%></option>
											<option value="11"><% If 1 = 2 Then %>Precio 
											oferta<% Else %><%=Replace(getsearch_OfertXLngStr("LtxtOfertPrice"), "{0}", txtOfert)%><% End If %></option>
											<option value="16"><%=getsearch_OfertXLngStr("LtxtRespPrice")%></option>
											</select> <select size="1" name="orden2">
											<option value="asc"><%=getsearch_OfertXLngStr("DtxtAsc")%></option>
											<option value="desc" selected>
											<%=getsearch_OfertXLngStr("DtxtDesc")%></option>
											</select></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td>&nbsp;</td>
											<td>&nbsp;</td>
											<td>
											<p align="right">
											<input type="submit" value="<%=getsearch_OfertXLngStr("DbtnSearch")%>" name="B1"></td>
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
<iframe id="ifGetValue" name="ifGetValue" style="display: none" height="169" width="84" src=""></iframe>
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