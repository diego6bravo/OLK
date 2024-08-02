<!--#include file="lang/listpenSearch.asp" -->
<%
If Request("getValType") = "" or Request("getValType") <> "" and InStr(Request("getValVal"), "*") = 0 and Request("getValType") <> "date" Then

LogNumFrom = Request("LogNumFrom")
LogNumTo = Request("LogNumTo")
dtFrom = Request("dtFrom")
dtTo = Request("dtTo")
Comments = Request("Comments")

setToVal = ""

SlpCodeFrom 	= ListPendSearchGetValue("SlpCodeFrom")
SlpCodeTo 		= ListPendSearchGetValue("SlpCodeTo")
CardCodeFrom 	= ListPendSearchGetValue("CardCodeFrom")
CardCodeTo 		= ListPendSearchGetValue("CardCodeTo")
CardNameFrom 	= ListPendSearchGetValue("CardNameFrom")
CardNameTo 		= ListPendSearchGetValue("CardNameTo")
GroupNameFrom 	= ListPendSearchGetValue("GroupNameFrom")
GroupNameTo 	= ListPendSearchGetValue("GroupNameTo")
CountryFrom 	= ListPendSearchGetValue("CountryFrom")
CountryTo 		= ListPendSearchGetValue("CountryTo")
ItemCodeFrom 	= ListPendSearchGetValue("ItemCodeFrom")
ItemCodeTo 		= ListPendSearchGetValue("ItemCodeTo")

Function ListPendSearchGetValue(Fld)
	If Request("getValType") = "" or Request("getValType") <> "" and Request("getValFld") <> Fld Then
		If setToVal <> "" and Request(Fld) = "" Then
			ListPendSearchGetValue = setToVal
		Else
			ListPendSearchGetValue = Request(Fld)
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
			ListPendSearchGetValue = NewValue
			
			If Right(Fld, 4) = "From" Then setToVal = NewValue
		Else
			ListPendSearchGetValue = ""
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
			alert('<%=getlistpenSearchLngStr("DtxtValNumVal")%>');
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
		document.frmSmallSearch.cmd.value = 'searchPend';
		document.frmSmallSearch.submit();
	}
}
function getDate(f)
{
	document.frmSmallSearch.getValType.value = 'date';
	document.frmSmallSearch.getValFld.value = f.name
	document.frmSmallSearch.getValVal.value = f.value;
	document.frmSmallSearch.cmd.value = 'searchPend';
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
					<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getlistpenSearchLngStr("LtxtOpenDocsSearch")%></font></b></td>
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
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtLogNum")%></font></td>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtFrom")%></font></td>
											<td>
											<input type="text" name="LogNumFrom" size="11" onchange="chkNum(this);if(document.frmSmallSearch.LogNumTo.value=='')document.frmSmallSearch.LogNumTo.value=this.value;" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=LogNumFrom%>'></td>
										</tr>
										<tr>
											<td width="79">&nbsp;</td>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtTo")%></font></td>
											<td>
											<input type="text" name="LogNumTo" size="11" onfocus="this.select();" onmouseup="event.preventDefault()" onchange="chkNum(this)" value='<%=LogNumTo%>'></td>
										</tr>
										<tr>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtDate")%></font></td>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtFrom")%></font></td>
											<td>
											
											<input readonly name="dtFrom" id="dtFrom" size="11" onclick="javascript:getDate(this);" value='<%=dtFrom%>'><img border="0" src="images/cal.gif" id="btnDtFrom" onclick="dtFrom.click()"></td>
										</tr>
										<tr>
											<td width="79">&nbsp;</td>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtTo")%></font></td>
											<td>
													<input readonly name="dtTo" id="dtTo" size="11" onclick="javascript:getDate(this);" value='<%=dtTo%>'><img border="0" src="images/cal.gif" id="btnDtTo" onclick="dtTo.click()"></td>
										</tr>
										<% If myAut.HasAuthorization(97) Then %>
										<tr>
											<td width="79"><font face="Verdana" size="1"><% If 1 = 2 Then %>Agents<% Else %><%=txtAgent%><% End If %></font></td>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtFrom")%></font></td>
											<td>
											<input type="text" name="SlpCodeFrom" size="15" maxlength="32" onchange="javascript:getValue('Slp', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=SlpCodeFrom%>'></td>
										</tr>
										<tr>
											<td width="79">&nbsp;</td>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtTo")%></font></td>
											<td>
											<input type="text" name="SlpCodeTo" size="15" maxlength="32" onchange="javascript:getValue('Slp', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=SlpCodeTo%>'></td>
										</tr>
										<% Else
										set rs = Server.CreateObject("ADODB.RecordSet") %>
										<input type="hidden" name="SlpCodeFrom" value="<%=mySession.GetAgentName%>">
										<input type="hidden" name="SlpCodeTo" value="<%=mySession.GetAgentName%>">
										<% End If %>
										<tr>
											<td width="79"><font face="Verdana" size="1"><% If 1 = 2 Then %><%=getlistpenSearchLngStr("DtxtClient")%> <% Else %><%=txtClient%><% End If %>- 
											<%=getlistpenSearchLngStr("DtxtCode")%></font></td>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtFrom")%></font></td>
											<td>
											<input type="text" name="CardCodeFrom" size="15" maxlength="15" onchange="javascript:getValue('Crd', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=CardCodeFrom%>'></td>
										</tr>
										<tr>
											<td width="79">&nbsp;</td>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtTo")%></font></td>
											<td>
											<input type="text" name="CardCodeTo" size="15" maxlength="15" onchange="javascript:getValue('Crd', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=CardCodeTo%>'></td>
										</tr>
										<tr>
											<td width="79"><font face="Verdana" size="1"><% If 1 = 2 Then %><%=getlistpenSearchLngStr("DtxtClient")%><% Else %><%=txtClient%><% End If %> 
											- <%=getlistpenSearchLngStr("DtxtName")%></font></td>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtFrom")%></font></td>
											<td>
											<input type="text" name="CardNameFrom" size="15" maxlength="15" onchange="javascript:getValue('CrdNam', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=CardNameFrom%>'></td>
										</tr>
										<tr>
											<td width="79">&nbsp;</td>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtTo")%></font></td>
											<td>
											<input type="text" name="CardNameTo" size="15"  maxlength="15" onchange="javascript:getValue('CrdNam', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=CardNameTo%>'></td>
										</tr>
										<tr>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtGroup")%></font></td>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtFrom")%></font></td>
											<td>
											<input type="text" name="GroupNameFrom" size="15" maxlength="20" onchange="javascript:getValue('Grp', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=GroupNameFrom%>'></td>
										</tr>
										<tr>
											<td width="79">&nbsp;</td>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtTo")%></font></td>
											<td>
											<input type="text" name="GroupNameTo" size="15" maxlength="20" onchange="javascript:getValue('Grp', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=GroupNameTo%>'></td>
										</tr>
										<tr>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtCountry")%></font></td>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtFrom")%></font></td>
											<td>
											<input type="text" name="CountryFrom" size="15" maxlength="100" onchange="javascript:getValue('Cty', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=CountryFrom%>'></td>
										</tr>
										<tr>
											<td width="79">&nbsp;</td>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtTo")%></font></td>
											<td>
											<input type="text" name="CountryTo" size="15" maxlength="100" onchange="javascript:getValue('Cty', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=CountryTo%>'></td>
										</tr>
										<tr>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtItem")%></font></td>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtFrom")%></font></td>
											<td>
											<input type="text" name="ItemCodeFrom" size="15" maxlength="20" onchange="javascript:getValue('Itm', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=ItemCodeFrom%>'></td>
										</tr>
										<tr>
											<td width="79">&nbsp;</td>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtTo")%></font></td>
											<td>
											<input type="text" name="ItemCodeTo" size="15" maxlength="20" onchange="javascript:getValue('Itm', this);" onfocus="this.select();" onmouseup="event.preventDefault()" value='<%=ItemCodeTo%>'></td>
										</tr>
										<tr>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtCommentaries")%></font></td>
											<td width="79">&nbsp;</td>
											<td>
											<textarea name="Comments" cols="16" rows="2" onfocus="this.select();" onmouseup="event.preventDefault()"><% If Comments <> "" Then %><%=Server.HTMLEncode(Comments)%><% End If %></textarea>
											</td>
										</tr>
										<tr>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtType")%></font></td>
											<td width="79">&nbsp;</td>
											<td>
											<select size="1" name="DocType">
									<option></option>
									<% If myApp.EnableOQUT Then %><option <% If Request("DocType") = "23" Then %>selected<% End If %> value="23"><%=txtQuotes%></option><% End If %>
									<% If myApp.EnableORDR Then %><option <% If Request("DocType") = "17" Then %>selected<% End If %> value="17"><%=txtOrdrs%></option><% End If %>
									</select></td>
										</tr>
										<tr>
											<td width="79"><font face="Verdana" size="1"><%=getlistpenSearchLngStr("DtxtOrder")%></font></td>
											<td width="79">&nbsp;</td>
											<td>
											<select size="1" name="orden1">
									<option value="0"><%=getlistpenSearchLngStr("DtxtLogNum")%></option>
									<option <% If Request("orden1") = "1" Then %>selected<% End If %> value="1"><%=getlistpenSearchLngStr("DtxtCode")%></option>
									<option <% If Request("orden1") = "2" Then %>selected<% End If %> value="2"><%=getlistpenSearchLngStr("DtxtName")%></option>
									<option <% If Request("orden1") = "3" or Request("orden1") = "" Then %>selected<% End If %> value="3"><%=getlistpenSearchLngStr("LtxtDocDate")%></option>
									<option <% If Request("orden1") = "7" Then %>selected<% End If %> value="7"><%=getlistpenSearchLngStr("LtxtDocType")%></option>
									<option <% If Request("orden1") = "9" Then %>selected<% End If %> value="9"><%=getlistpenSearchLngStr("DtxtTotal")%></option>
									</select></td>
										</tr>
										<tr>
											<td width="79">&nbsp;</td>
											<td width="79">&nbsp;</td>
											<td>
											<select size="1" name="orden2">
									<option value="A"><%=getlistpenSearchLngStr("DtxtAsc")%></option>
									<option value="D" selected><%=getlistpenSearchLngStr("DtxtDesc")%></option>
									</select></td>
										</tr>										<tr>
											<td colspan="3">
											<p align="center">
											<input type="button" value="<%=getlistpenSearchLngStr("DtxtClear")%>" name="btnClear" onclick="javascript:window.location.href='?cmd=searchPend';"></td>
										</tr>
										<tr>
											<td colspan="3">
											<p align="center">
											<input type="submit" value="<%=getlistpenSearchLngStr("DbtnSearch")%>" name="btnSearch" onclick="javascript:document.frmSmallSearch.cmd.value='pendientes';"></td>
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

<div align="center">
	<table border="0" cellpadding="0">
		<tr>
			<td valign="top">
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getlistpenSearchLngStr("LtxtOpenDocsSearch")%> 
					- <%=getlistpenSearchLngStr("LtxtSelVal")%></font></b></td>
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
										Dim getValSelect
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
										
										monthNames = array("", getlistpenSearchLngStr("DtxtMonthJanuary"), getlistpenSearchLngStr("DtxtMonthFebruary"), getlistpenSearchLngStr("DtxtMonthMarch"), getlistpenSearchLngStr("DtxtMonthApril"), getlistpenSearchLngStr("DtxtMonthMay"), getlistpenSearchLngStr("DtxtMonthJune"), getlistpenSearchLngStr("DtxtMonthJuly"), getlistpenSearchLngStr("DtxtMonthAugust"), getlistpenSearchLngStr("DtxtMonthSeptember"), getlistpenSearchLngStr("DtxtMonthOctober"), getlistpenSearchLngStr("DtxtMonthNovember"), getlistpenSearchLngStr("DtxtMonthDecember"))
										dayNames = array("", getlistpenSearchLngStr("DtxtSmallDayMonday"), getlistpenSearchLngStr("DtxtSmallDayTuesday"), getlistpenSearchLngStr("DtxtSmallDayWednesday"), getlistpenSearchLngStr("DtxtSmallDayThursday"), getlistpenSearchLngStr("DtxtSmallDayFriday"), getlistpenSearchLngStr("DtxtSmallDaySaturday"), getlistpenSearchLngStr("DtxtSmallDaySunday")) 
										txtToday = getlistpenSearchLngStr("DtxtToday")
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
												<% strDate = myApp.DateFormat
												strDate = Replace(strDate, "yyyy", selYear)
												strDate = Replace(strDate, "MM", Right("0" & selMonth, 2))
												strDate = Replace(strDate, "dd", Right("0" & selDay, 2)) %>
												<input type="button" name="btnSubmit" value="<%=getlistpenSearchLngStr("DtxtAccept")%>" onclick="javascript:setSmallSearchVal('<%=strDate%>');"></td>
												<td>
												<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
												<input type="button" name="btnCancel" value="<%=getlistpenSearchLngStr("DtxtCancel")%>" onclick="javascript:cancelSmallSearchVal();"></td>
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
	document.frmSmallSearch.cmd.value = 'searchPend';
	document.frmSmallSearch.submit();
}
function cancelSmallSearchVal()
{
	document.frmSmallSearch.<%=Request.Form("getValFld")%>.value = '';
	document.frmSmallSearch.cmd.value = 'searchPend';
	document.frmSmallSearch.submit();
}
</script>
<% End If %>