<!--#include file="lang/openActivitiesSearch.asp" -->
<%
If Request("getValType") = "" or Request("getValType") <> "" and InStr(Request("getValVal"), "*") = 0 and Request("getValType") <> "date" Then

LogNumFrom = Request("LogNumFrom")
LogNumTo = Request("LogNumTo")
dtFrom = Request("dtFrom")
dtTo = Request("dtTo")
Comments = Request("Comments")

setToVal = ""

SlpCodeFrom 	= OpenActivitiesSearchGetValue("SlpCodeFrom")
SlpCodeTo 		= OpenActivitiesSearchGetValue("SlpCodeTo")
CardCodeFrom 	= OpenActivitiesSearchGetValue("CardCodeFrom")
CardCodeTo 		= OpenActivitiesSearchGetValue("CardCodeTo")
CardNameFrom 	= OpenActivitiesSearchGetValue("CardNameFrom")
CardNameTo 		= OpenActivitiesSearchGetValue("CardNameTo")
GroupNameFrom 	= OpenActivitiesSearchGetValue("GroupNameFrom")
GroupNameTo 	= OpenActivitiesSearchGetValue("GroupNameTo")
CountryFrom 	= OpenActivitiesSearchGetValue("CountryFrom")
CountryTo 		= OpenActivitiesSearchGetValue("CountryTo")

Function OpenActivitiesSearchGetValue(Fld)
	If Request("getValType") = "" or Request("getValType") <> "" and Request("getValFld") <> Fld Then
		If setToVal <> "" and Request(Fld) = "" Then
			OpenActivitiesSearchGetValue = setToVal
		Else
			OpenActivitiesSearchGetValue = Request(Fld)
		End If
		setToVal = ""
	Else
		
		Dim getVal
		set getVal = new clsGetValValue
		getVal.ValueType = Request("getValType")
		getVal.ValueField = Request("getValFld")
		getVal.Value = Request("getValVal")
		
		NewValue = getVal.GetValue
		
		If NewValue <> "" Then
			OpenActivitiesSearchGetValue = NewValue
			
			If Right(Fld, 4) = "From" Then setToVal = NewValue
		Else
			OpenActivitiesSearchGetValue = ""
		End If 
	End If
End Function

%>
<script type="text/javascript">
function chkNum(fld)
{
	if (fld.value != '')
	{
		if (isNaN(parseInt(fld.value)))
		{
			alert('<%=getopenActivitiesSearchLngStr("DtxtValNumVal")%>');
			fld.value = '';
			fld.focus();
		}
	}
}
function getValue(t, f)
{
	if (f.value != '')
	{
		document.frmSmallSearch.getValType.value = t;
		document.frmSmallSearch.getValFld.value = f.name;
		document.frmSmallSearch.getValVal.value = f.value;
		document.frmSmallSearch.cmd.value = 'openActivitiesSearch';
		document.frmSmallSearch.submit();
	}
}
function getDate(f)
{
	document.frmSmallSearch.getValType.value = 'date';
	document.frmSmallSearch.getValFld.value = f.name
	document.frmSmallSearch.getValVal.value = f.value;
	document.frmSmallSearch.cmd.value = 'openActivitiesSearch';
	document.frmSmallSearch.submit();
}
</script>
<form method="POST" action="operaciones.asp" name="frmSmallSearch">
<input type="hidden" name="cmd" value="pendientes">
<input type="hidden" name="getValType" value="">
<input type="hidden" name="getValFld" value="">
<input type="hidden" name="getValVal" value="">
<div align="center">
	<table border="0" cellpadding="0">
		<tr>
			<td valign="top">
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("LtxtOpenActivitiesSea")%></font></b></td>
				</tr>
				<tr>
					<td>
					<table border="0" cellpadding="0" width="100%" cellspacing="0">
						<tr>
							<td>
							<table border="0" cellpadding="0" width="100%" cellspacing="1">
								<tr>
									<td>
									<table border="0" cellpadding="0" width="100%" cellspacing="1">
										<tr>
											<td width="79"><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("DtxtSource")%></font></td>
											<td width="79">&nbsp;</td>
											<td>
											<select name="cmbSourceType" size="1">
											<option value=""><%=getopenActivitiesSearchLngStr("DtxtAll")%></option>
											<option value="O"><%=getopenActivitiesSearchLngStr("DtxtOLK")%></option>
											<option value="S"><%=getopenActivitiesSearchLngStr("DtxtSAP")%></option>
											</select></td>
										</tr>
										<tr>
											<td width="79"><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("DtxtLogNum")%></font></td>
											<td width="79"><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("DtxtFrom")%></font></td>
											<td>
											<input type="text" name="LogNumFrom" size="11" onchange="chkNum(this);if(document.frmSmallSearch.LogNumTo.value=='')document.frmSmallSearch.LogNumTo.value=this.value;" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=LogNumFrom%>'></td>
										</tr>
										<tr>
											<td width="79">&nbsp;</td>
											<td width="79"><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("DtxtTo")%></font></td>
											<td>
											<input type="text" name="LogNumTo" size="11" onfocus="this.select();" onmouseup="event.preventDefault()" onchange="chkNum(this)" value='<%=LogNumTo%>'></td>
										</tr>
										<tr>
											<td width="79"><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("DtxtDate")%></font></td>
											<td width="79"><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("DtxtFrom")%></font></td>
											<td>
											
											<input readonly name="dtFrom" id="dtFrom" size="11" onclick="javascript:getDate(this);" value='<%=dtFrom%>'><img border="0" src="images/cal.gif" id="btnDtFrom" onclick="dtFrom.click()"></td>
										</tr>
										<tr>
											<td width="79">&nbsp;</td>
											<td width="79"><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("DtxtTo")%></font></td>
											<td>
													<input readonly name="dtTo" id="dtTo" size="11" onclick="javascript:getDate(this);" value='<%=dtTo%>'><img border="0" src="images/cal.gif" id="btnDtTo" onclick="dtTo.click()"></td>
										</tr>
										<% If myAut.HasAuthorization(97) Then %>
										<tr>
											<td width="79"><font face="Verdana" size="1"><% If 1 = 2 Then %>Agents<% Else %><%=txtAgent%><% End If %></font></td>
											<td width="79"><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("DtxtFrom")%></font></td>
											<td>
											<input type="text" name="SlpCodeFrom" size="15" maxlength="32" onchange="javascript:getValue('Slp', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=SlpCodeFrom%>'></td>
										</tr>
										<tr>
											<td width="79">&nbsp;</td>
											<td width="79"><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("DtxtTo")%></font></td>
											<td>
											<input type="text" name="SlpCodeTo" size="15" maxlength="32" onchange="javascript:getValue('Slp', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=SlpCodeTo%>'></td>
										</tr>
										<% Else
										set rs = Server.CreateObject("ADODB.RecordSet") %>
										<input type="hidden" name="SlpCodeFrom" value="<%=mySession.GetAgentName%>">
										<input type="hidden" name="SlpCodeTo" value="<%=mySession.GetAgentName%>">
										<% End If %>
										<tr>
											<td width="79"><font face="Verdana" size="1"><% If 1 = 2 Then %>|D:txtClient| <% Else %><%=txtClient%><% End If %>- 
											<%=getopenActivitiesSearchLngStr("DtxtCode")%></font></td>
											<td width="79"><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("DtxtFrom")%></font></td>
											<td>
											<input type="text" name="CardCodeFrom" size="15" maxlength="15" onchange="javascript:getValue('Crd', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=CardCodeFrom%>'></td>
										</tr>
										<tr>
											<td width="79">&nbsp;</td>
											<td width="79"><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("DtxtTo")%></font></td>
											<td>
											<input type="text" name="CardCodeTo" size="15" maxlength="15" onchange="javascript:getValue('Crd', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=CardCodeTo%>'></td>
										</tr>
										<tr>
											<td width="79"><font face="Verdana" size="1"><% If 1 = 2 Then %>|D:txtClient|<% Else %><%=txtClient%><% End If %> 
											- <%=getopenActivitiesSearchLngStr("DtxtName")%></font></td>
											<td width="79"><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("DtxtFrom")%></font></td>
											<td>
											<input type="text" name="CardNameFrom" size="15" maxlength="15" onchange="javascript:getValue('CrdNam', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=CardNameFrom%>'></td>
										</tr>
										<tr>
											<td width="79">&nbsp;</td>
											<td width="79"><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("DtxtTo")%></font></td>
											<td>
											<input type="text" name="CardNameTo" size="15"  maxlength="15" onchange="javascript:getValue('CrdNam', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=CardNameTo%>'></td>
										</tr>
										<tr>
											<td width="79"><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("DtxtGroup")%></font></td>
											<td width="79"><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("DtxtFrom")%></font></td>
											<td>
											<input type="text" name="GroupNameFrom" size="15" maxlength="20" onchange="javascript:getValue('Grp', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=GroupNameFrom%>'></td>
										</tr>
										<tr>
											<td width="79">&nbsp;</td>
											<td width="79"><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("DtxtTo")%></font></td>
											<td>
											<input type="text" name="GroupNameTo" size="15" maxlength="20" onchange="javascript:getValue('Grp', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=GroupNameTo%>'></td>
										</tr>
										<tr>
											<td width="79"><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("DtxtCountry")%></font></td>
											<td width="79"><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("DtxtFrom")%></font></td>
											<td>
											<input type="text" name="CountryFrom" size="15" maxlength="100" onchange="javascript:getValue('Cty', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=CountryFrom%>'></td>
										</tr>
										<tr>
											<td width="79">&nbsp;</td>
											<td width="79"><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("DtxtTo")%></font></td>
											<td>
											<input type="text" name="CountryTo" size="15" maxlength="100" onchange="javascript:getValue('Cty', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=CountryTo%>'></td>
										</tr>
										<tr>
											<td width="79"><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("DtxtActivity")%></font></td>
											<td width="79">&nbsp;</td>
											<td>
											<select size="1" name="DocType">
										<option></option>
										<option <% If Request("DocType") = "C" Then %>selected<% End If %> value="C"><%=getopenActivitiesSearchLngStr("DtxtConv")%></option>
										<option <% If Request("DocType") = "M" Then %>selected<% End If %> value="M"><%=getopenActivitiesSearchLngStr("DtxtMeeting")%></option>
										<option <% If Request("DocType") = "E" Then %>selected<% End If %> value="E"><%=getopenActivitiesSearchLngStr("DtxtNote")%></option>
										<option <% If Request("DocType") = "O" Then %>selected<% End If %> value="O"><%=getopenActivitiesSearchLngStr("DtxtOther")%></option>
										<option <% If Request("DocType") = "T" Then %>selected<% End If %> value="T"><%=getopenActivitiesSearchLngStr("DtxtTask")%></option>
										</select></td>
										</tr>
										<tr>
											<td width="79"><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("DtxtOrder")%></font></td>
											<td width="79">&nbsp;</td>
											<td>
											<select size="1" name="orden1">
											<option value="1"><%=getopenActivitiesSearchLngStr("DtxtLogNum")%></option>
											<option <% If Request("orden1") = "T1.CardCode" Then %>selected<% End If %> value="T1.CardCode"><%=getopenActivitiesSearchLngStr("DtxtCode")%></option>
											<option <% If Request("orden1") = "CardName" Then %>selected<% End If %> value="CardName"><%=getopenActivitiesSearchLngStr("DtxtName")%></option>
											<option <% If Request("orden1") = "2" or Request("orden1") = "" Then %>selected<% End If %> value="2"><%=getopenActivitiesSearchLngStr("LtxtCntDate")%></option>
											<option <% If Request("orden1") = "Action" Then %>selected<% End If %> value="Action"><%=getopenActivitiesSearchLngStr("DtxtActivity")%></option>
											</select></td>
										</tr>
										<tr>
											<td width="79">&nbsp;</td>
											<td width="79">&nbsp;</td>
											<td>
											<select size="1" name="orden2">
									<option value="asc"><%=getopenActivitiesSearchLngStr("DtxtAsc")%></option>
									<option value="desc" <% If Request("orden2") = "" or Request("orden2") = "desc" Then %>selected<% End If %>><%=getopenActivitiesSearchLngStr("DtxtDesc")%></option>
									</select></td>
										</tr>										<tr>
											<td colspan="3">
											<p align="center">
											<input type="button" value="<%=getopenActivitiesSearchLngStr("DtxtClear")%>" name="btnClear" onclick="javascript:window.location.href='?cmd=openActivitiesSearch';"></td>
										</tr>
										<tr>
											<td colspan="3">
											<p align="center">
											<input type="submit" value="<%=getopenActivitiesSearchLngStr("DbtnSearch")%>" name="btnSearch" onclick="javascript:document.frmSmallSearch.cmd.value='openActivities';"></td>
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
<% Else %>
<link rel="stylesheet" href="Reportes/style.css">

<div align="center">
	<table border="0" cellpadding="0">
		<tr>
			<td valign="top">
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getopenActivitiesSearchLngStr("LtxtOpenActivitiesSea")%> 
					- <%=getopenActivitiesSearchLngStr("LtxtSelVal")%></font></b></td>
				</tr>
				<tr>
					<td>
					<table border="0" cellpadding="0" width="100%" cellspacing="0">
						<tr>
							<td>
							<table border="0" cellpadding="0" width="100%" cellspacing="1">
								<tr>
									<td>
									<% 
									If Request("getValType") <> "date" Then
										'Dim getValSelect
										set getValSelect = New clsGetValValueSelect
										getValSelect.ValueType = Request("getValType")
										getValSelect.ValueField = Request("getValFld")
										getValSelect.Value = Request("getValVal")
										getValSelect.OnClick = "javascript:setSmallSearchVal('{0}');"
										getValSelect.OnCancel = "cancelSmallSearchVal();"
										getValSelect.ShowValues
									Else 
										Response.Buffer = true
										Response.Addheader "Pragma","no-cache"
										'============================================================================================
										' VARIABLES
										'============================================================================================
										'Dim i
										scriptName = Request.ServerVariables("SCRIPT_NAME")
										
										imgPath = "images/"
										
										monthNames = array("", getopenActivitiesSearchLngStr("DtxtMonthJanuary"), getopenActivitiesSearchLngStr("DtxtMonthFebruary"), getopenActivitiesSearchLngStr("DtxtMonthMarch"), getopenActivitiesSearchLngStr("DtxtMonthApril"), getopenActivitiesSearchLngStr("DtxtMonthMay"), getopenActivitiesSearchLngStr("DtxtMonthJune"), getopenActivitiesSearchLngStr("DtxtMonthJuly"), getopenActivitiesSearchLngStr("DtxtMonthAugust"), getopenActivitiesSearchLngStr("DtxtMonthSeptember"), getopenActivitiesSearchLngStr("DtxtMonthOctober"), getopenActivitiesSearchLngStr("DtxtMonthNovember"), getopenActivitiesSearchLngStr("DtxtMonthDecember"))
										dayNames = array("", getopenActivitiesSearchLngStr("DtxtSmallDayMonday"), getopenActivitiesSearchLngStr("DtxtSmallDayTuesday"), getopenActivitiesSearchLngStr("DtxtSmallDayWednesday"), getopenActivitiesSearchLngStr("DtxtSmallDayThursday"), getopenActivitiesSearchLngStr("DtxtSmallDayFriday"), getopenActivitiesSearchLngStr("DtxtSmallDaySaturday"), getopenActivitiesSearchLngStr("DtxtSmallDaySunday")) 
										txtToday = getopenActivitiesSearchLngStr("DtxtToday")
										%>
										<table border="0" cellspacing="0" width="100%">
											<form method="POST" name="frmSmallSearchDate" action="operaciones.asp">
											<tr class="TblAfueraMnu">
												<td colspan="2" align="center">
												<% 
												If Request("d") = "" Then
													calVal = Request(Request("getValFld"))
													If calVal <> "" Then
														calVald = Mid(calVal, InStr(myApp.DateFormat, "dd"), 2)
														calValm = Mid(calVal, InStr(myApp.DateFormat, "MM"), 2)
														calValy = Mid(calVal, InStr(myApp.DateFormat, "yyyy"), 4)
													End If
												Else
														calVald = Request("d")
														calValm = Request("m")
														calValy = Request("y")
												End If %>
												<%
												' call calendar
												makeCalendar calVald,calValm,calValy,""
												%>
												</td>
											</tr>
											<tr class="TblAfueraMnu">
												<td colspan="2">
												<p align="center">
												</td>
											</tr>
												<% 	For each itm in Request.Form
												If itm <> "d" and itm <> "m" and itm <> "y" and itm <> "l" and itm <> "s" Then %>
												<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
												<% 	End If
												Next %>
											<tr class="TblAfueraMnu">
												<td>
												<p align="left">
												<% 
												strDate = myApp.DateFormat
												strDate = Replace(strDate, "yyyy", selYear)
												strDate = Replace(strDate, "MM", Right("0" & selMonth, 2))
												strDate = Replace(strDate, "dd", Right("0" & selDay, 2)) %>
												<input type="button" name="btnSubmit" value="<%=getopenActivitiesSearchLngStr("DtxtAccept")%>" onclick="javascript:setSmallSearchVal('<%=strDate%>');"></td>
												<td>
												<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
												<input type="button" name="btnCancel" value="<%=getopenActivitiesSearchLngStr("DtxtCancel")%>" onclick="javascript:cancelSmallSearchVal();"></td>
											</tr>
											</form>
										</table>
										<%

									End If %>
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
<form name="frmSmallSearch" action="operaciones.asp" method="post">
<% For each itm in Request.Form
	If itm <> "getValType" and itm <> "getValFld" and itm <> "getValVal" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request.Form(itm)%>">
<% End If
Next %>
</form>
<script language="javascript">
function setSmallSearchVal(value)
{
	document.frmSmallSearch.<%=Request.Form("getValFld")%>.value = value;
	<% If Right(Request.Form("getValFld"), 4) = "From" Then %>
	var fldTo = document.frmSmallSearch.<%=Left(Request.Form("getValFld"), Len(Request.Form("getValFld"))-4)%>To;
	if (fldTo.value == '') fldTo.value = value;
	<% End If %>
	document.frmSmallSearch.cmd.value = 'openActivitiesSearch';
	document.frmSmallSearch.submit();
}
function cancelSmallSearchVal()
{
	document.frmSmallSearch.<%=Request.Form("getValFld")%>.value = '';
	document.frmSmallSearch.cmd.value = 'openActivitiesSearch';
	document.frmSmallSearch.submit();
}
</script>
<% End If %>