<% addLngPathStr = "ventas/" %>
<!--#include file="lang/search_ventasX.asp" -->
<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>
<script language="javascript">
var retAcct;
function Start(page, retAction) 
{
	retAcct = retAction;
	OpenWin = this.open(page, "DatePicker", "toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no, width=240,height=220");
}

function setTimeStamp(retAction, varDate) 
{
	retAcct.value = varDate;
}

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
	if (updFld.name.substring(0, 4) == "From")
	{
		setFldName = updFld.name.subsring(0, updFld.name.length-4);
		fldTo = document.getElementById(setFldName + 'To');
		if (fldTo.value == '') { fldTo.value = value; fldTo.select(); }
	}
}
</script>
<form method="post" target="ifGetValue" name="frmGetValue" action="topGetValue.asp">
<input type="hidden" name="Type" value="">
<input type="hidden" name="searchStr" value="">
</form>
<form method="POST" action="searchOpenedDocs.asp" name="frmSmallSearch">
<div align="center">
	<table border="0" cellpadding="0" width="499">
		<tr>
			<td>
			<p align="center">
			<img border="0" src="design/0/images/search_top.jpg" width="407" height="140"></td>
		</tr>
		<tr>
			<td valign="top">
			<table border="0" cellpadding="0" width="100%">
				<tr class="GeneralTlt">
					<td>
					<p align="left"><%=getsearch_ventasXLngStr("LttlPendDocSearch")%></td>
				</tr>
				<tr>
					<td>
					<table border="0" cellpadding="0" width="100%" cellspacing="0">
						<tr class="GeneralTlt">
							<td>
							<table border="0" cellpadding="0" width="100%" cellspacing="1">
								<tr class="GeneralTbl">
									<td>
									<table border="0" cellpadding="0" width="100%" cellspacing="1">
										<tr class="GeneralTblBold2">
											<td width="79"><%=getsearch_ventasXLngStr("DtxtLogNum")%></td>
											<td>
											<input type="text" name="LogNumFrom" size="11" onkeydown="return valKeyNum(event);" onchange="if(document.frmSmallSearch.LogNumTo.value=='')document.frmSmallSearch.LogNumTo.value=this.value;" onfocus="this.select();" value='<%=Request("LogNumFrom")%>'> 
											-
											<input type="text" name="LogNumTo" size="11" onfocus="this.select();" onkeydown="return valKeyNum(event);" value='<%=Request("LogNumTo")%>'></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td width="79"><%=getsearch_ventasXLngStr("DtxtDate")%></td>
											<td>
											<table cellpadding="0" cellspacing="0" border="0">
												<tr>
													<td>
													<input readonly name="dtFrom" size="11" onclick="btnDtFrom.click()" value='<%=Request("dtFrom")%>'></td>
													<td>&nbsp;<img border="0" src="images/cal.gif" id="btnDtFrom"></td>
													<td>&nbsp;-&nbsp;</td>
													<td>
													<input readonly name="dtTo" size="11" onclick="btnDtTo.click()" value='<%=Request("dtTo")%>'></td>
													<td>&nbsp;<img border="0" src="images/cal.gif" id="btnDtTo"></td>
												</tr>
											</table>
											</td>
										</tr>
										<% If myAut.HasAuthorization(97) Then %>
										<tr class="GeneralTblBold2">
											<td width="79"><% If 1 = 2 Then %>Agents<% Else %><%=txtAgent%><% End If %></td>
											<td>
											<input type="text" name="SlpCodeFrom" size="15" onkeydown="return chkMax(event, this, 32);" onchange="javascript:getValue('Slp', this);" onfocus="this.select();" value='<%=Request("SlpCodeFrom")%>'> 
											-
											<input type="text" name="SlpCodeTo" size="15" onkeydown="return chkMax(event, this, 32);" onchange="javascript:getValue('Slp', this);" onfocus="this.select();" value='<%=Request("SlpCodeTo")%>'></td>
										</tr>
										<% Else %>
										<input type="hidden" name="SlpCodeFrom" value="<%=myHTMLEncode(AgentName)%>">
										<input type="hidden" name="SlpCodeTo" value="<%=myHTMLEncode(AgentName)%>">
										<% End If %>
										<tr class="GeneralTblBold2">
											<td width="79"><% If 1 = 2 Then %><%=getsearch_ventasXLngStr("DtxtClient")%><% Else %><%=txtClient%><% End If %></td>
											<td>
											<input type="text" name="CardCodeFrom" size="15" onkeydown="return chkMax(event, this, 15);" onchange="javascript:getValue('Crd', this);" onfocus="this.select();" value='<%=Request("CardCodeFrom")%>'> 
											-
											<input type="text" name="CardCodeTo" size="15" onkeydown="return chkMax(event, this, 15);" onchange="javascript:getValue('Crd', this);" onfocus="this.select();" value='<%=Request("CardCodeTo")%>'></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td width="79"><%=getsearch_ventasXLngStr("DtxtGroup")%></td>
											<td>
											<input type="text" name="GroupNameFrom" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('Grp', this);" onfocus="this.select();" value='<%=Request("GroupNameFrom")%>'> 
											-
											<input type="text" name="GroupNameTo" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('Grp', this);" onfocus="this.select();" value='<%=Request("GroupNameTo")%>'></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td width="79"><%=getsearch_ventasXLngStr("DtxtCountry")%></td>
											<td>
											<input type="text" name="CountryFrom" size="20" onkeydown="return chkMax(event, this, 100);" onchange="javascript:getValue('Cty', this);" onfocus="this.select();" value='<%=Request("CountryFrom")%>'> 
											-
											<input type="text" name="CountryTo" size="20" onkeydown="return chkMax(event, this, 100);" onchange="javascript:getValue('Cty', this);" onfocus="this.select();" value='<%=Request("CountryTo")%>'></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td width="79"><%=getsearch_ventasXLngStr("DtxtItem")%></td>
											<td>
											<input type="text" name="ItemCodeFrom" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('Itm', this);" onfocus="this.select();" value='<%=Request("ItemCodeFrom")%>'> 
											-
											<input type="text" name="ItemCodeTo" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('Itm', this);" onfocus="this.select();" value='<%=Request("ItemCodeTo")%>'></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td width="79"><%=getsearch_ventasXLngStr("DtxtCommentaries")%></td>
											<td>
											<input type="text" name="Comments" size="20" onfocus="this.select();" value='<%=Request("Comments")%>' style="width: 90%"> 
											</td>
										</tr>
										<tr class="GeneralTblBold2">
											<td width="79"><%=getsearch_ventasXLngStr("DtxtType")%></td>
											<td>
											<select size="1" name="DocType">
									<option></option>
									<% If myApp.EnableOQUT Then %><option <% If Request("DocType") = "23" Then %>selected<% End If %> value="23"><%=txtQuotes%></option><% End If %>
									<% If myApp.EnableORDR Then %><option <% If Request("DocType") = "17" Then %>selected<% End If %> value="17"><%=txtOrdrs%></option><% End If %>
									<% If myApp.EnableODLN Then %><option <% If Request("DocType") = "15" Then %>selected<% End If %> value="15"><%=txtOdlns%></option><% End If %>
									<% If myApp.EnableODPIReq Then %><option <% If Request("DocType") = "203" Then %>selected<% End If %> value="203"><%=txtODPIReqs%></option><% End If %>
									<% If myApp.EnableODPIInv Then %><option <% If Request("DocType") = "204" Then %>selected<% End If %> value="204"><%=txtODPIInvs%></option><% End If %>
									<% If myApp.EnableOINV Then %><option <% If Request("DocType") = "13" Then %>selected<% End If %> value="13"><%=txtInvs%></option><% End If %>
									<% If myApp.EnableOINVRes Then %><option <% If Request("DocType") = "-13" Then %>selected<% End If %> value="-13"><%=txtInvsRes%></option><% End If %>
									<% If myApp.EnableCashInv Then %><option <% If Request("DocType") = "48" Then %>selected<% End If %> value="48"><%=txtInvs%>/<%=txtRct%></option><% End If %>
									<% If myApp.EnableORCT Then %><option <% If Request("DocType") = "24" Then %>selected<% End If %> value="24"><%=txtRcts%></option><% End If %>
									</select></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td width="79"><%=getsearch_ventasXLngStr("DtxtOrder")%></td>
											<td>
											<select size="1" name="orden1">
									<option value="0"><%=getsearch_ventasXLngStr("DtxtLogNum")%></option>
									<option <% If Request("orden1") = "4" Then %>selected<% End If %> value="4"><%=txtAgent%></option>
									<option <% If Request("orden1") = "1" Then %>selected<% End If %> value="1"><%=getsearch_ventasXLngStr("DtxtCode")%></option>
									<option <% If Request("orden1") = "2" Then %>selected<% End If %> value="2"><%=getsearch_ventasXLngStr("DtxtName")%></option>
									<option <% If Request("orden1") = "5" Then %>selected<% End If %> value="5"><%=getsearch_ventasXLngStr("DtxtGroup")%></option>
									<option <% If Request("orden1") = "6" Then %>selected<% End If %> value="6"><%=getsearch_ventasXLngStr("DtxtCountry")%></option>
									<option <% If Request("orden1") = "3" Then %>selected<% End If %> value="3"><%=getsearch_ventasXLngStr("LtxtDocDate")%></option>
									<option <% If Request("orden1") = "7" Then %>selected<% End If %> value="7"><%=getsearch_ventasXLngStr("LtxtDocType")%></option>
									<option <% If Request("orden1") = "8" Then %>selected<% End If %> value="8"><%=getsearch_ventasXLngStr("DtxtState")%></option>
									<option <% If Request("orden1") = "9" Then %>selected<% End If %> value="9"><%=getsearch_ventasXLngStr("DtxtTotal")%></option>
									</select> <select size="1" name="orden2">
									<option value="A"><%=getsearch_ventasXLngStr("DtxtAsc")%></option>
									<option value="D" selected><%=getsearch_ventasXLngStr("DtxtDesc")%></option>
									</select></td>
										</tr>
										<tr class="GeneralTblBold2">
											<td colspan="2">
											<p align="center">
											<input type="submit" value="<%=getsearch_ventasXLngStr("DbtnSearch")%>" name="B1"></td>
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