<!--#include file="top.asp" -->
<!-- #INCLUDE file="FCKeditor/fckeditor.asp" -->
<!--#include file="lang/adminDocConf.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<!--#include file="accountControl.asp"-->  
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	background-color: #E2F3FC;
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
.style2 {
	background-color: #F7FBFF;
}
</style>
</head>
<%
conn.execute("use [" & Session("olkdb") & "]")
If Request("Object") <> "" Then obj = CLng(Request("Object")) Else obj = -1

If Request("Save") and obj <> -1 Then
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKAdminDocConf" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@ObjID") = obj
	If Request("Active") = "Y" Then cmd("@Active") = "Y"
	If Request("ActiveClient") = "Y" Then cmd("@ActiveClient") = "Y"
	If Request("Confirm") = "Y" Then cmd("@Confirm") = "Y"
	If Request("ViewObjPrint") = "Y" Then cmd("@ViewObjPrint") = "Y"
	If Request("ViewObjPrintClient") = "Y" Then cmd("@ViewObjPrintClient") = "Y"
	
	If obj = 48 Then
		If Request("Series2") <> "" Then cmd("@Series2") = Request("Series2")
		If Request("Series2Client") <> "" Then cmd("@Series2Client") = Request("Series2Client")
	End If
	
	If obj = 15 or obj = 17 or obj = 23 or obj = 48 Then
		If Request("SeriesClient") <> "" Then cmd("@SeriesClient") = Request("SeriesClient")
	End If
	
	If obj = 24 Then
		If Request("PrintContact") = "Y" Then PrintContact = "Y" Else PrintContact = "N"
		If Request("PrintAddress") = "Y" Then PrintAddress = "Y" Else PrintAddress = "N"
		If Request("ORCTContraComp") = "Y" Then ORCTContraComp = "Y" Else ORCTContraComp = "N"
		If Request("ApplyOpenRctToInvBal") = "Y" Then ApplyOpenRctToInvBal = "Y" Else ApplyOpenRctToInvBal = "N"
		
		cmd("@PrintContact") = PrintContact
		cmd("@PrintAddress") = PrintAddress
		cmd("@ORCTContraComp") = ORCTContraComp
		cmd("@ApplyOpenRctToInvBal") = ApplyOpenRctToInvBal
	End If
	If obj = 24 or obj = 48 Then
		If Request("PrintPaidSum") = "Y" Then PrintPaidSum = "Y" Else PrintPaidSum = "N"
		If Request("IgnoreSystemChecksFilter") = "Y" Then IgnoreSystemChecksFilter = "Y" Else IgnoreSystemChecksFilter = "N"
		
		cmd("@PrintPaidSum") = PrintPaidSum
		cmd("@IgnoreSystemChecksFilter") = IgnoreSystemChecksFilter
		
		If Request("CashAcct") <> "" Then cmd("@CashAcct") = Request("CashAcct")
		If Request("CheckAcct") <> "" Then cmd("@CheckAcct") = Request("CheckAcct")
		If Request("ChecksFilter") <> "" Then cmd("@ChecksFilter") = Request("ChecksFilter")
	End If
	
	If obj = 17 Then
		If Request("Verfy3dx") <> "" Then cmd("@Verfy3dx") = Request("Verfy3dx")
		cmd("@VerfyBtch") = Request("VerfyBtch")
	End If
	
	If Request("PrintCmpPaper") = "Y" Then cmd("@PrintCmpPaper") = "Y" 
	If Request("Series") <> "" Then cmd("@Series") = Request("Series")
	If Request("txtNote") <> "" Then cmd("@Note") = Request("txtNote")
	
	If obj = 13 Then
		If Request("EnResInv") = "Y" Then EnResInv = "Y" Else EnResInv = "N"
		If Request("DefResInv") = "Y" Then DefResInv = "Y" Else DefResInv = "N"
		cmd("@EnResInv") = EnResInv
		cmd("@DefResInv") = DefResInv
	End If
	
	If obj = 48 Then
		If Request("ClientReservedInvoice") = "Y" Then ClientReservedInvoice = "Y" Else ClientReservedInvoice = "N"
		cmd("@ClientReservedInvoice") = ClientReservedInvoice
	End If
	
	If obj = 2 or obj = 4 Then
		If Request("ChkAutGen") = "Y" Then ChkAutGen = "Y" Else ChkAutGen = "N"
		cmd("@EnableAutoGenCode") = ChkAutGen
		If Request("AutoGenQry") <> "" Then cmd("@AutoGenQry") = Request("AutoGenQry")
	End If
	
	cmd.execute()

	If obj = 17 Then myApp.LoadVerfyOrders
	myApp.LoadAdminObjConf
	If obj = 24 or obj = 48 Then myApp.LoadAdminDocConf
	myApp.ResetLastUpdate
	If obj = 2 or obj = 4 Then 
		myApp.LoadAutoGen
		Response.Redirect "adminDocConfUpdate.aspx?obj=" & obj & "&dbID=" & Session("ID") & "&dbName=" & Session("olkdb")
	End If
End If
%>
<script language="javascript">
function Start2(theURL, popW, popH, type) { // V 1.0
var winleft = (screen.width - popW) / 2;
var winUp = (screen.height - popH) / 2;
winProp = 'width='+popW+',height='+popH+',left='+winleft+',top='+winUp+',toolbar=no,scrollbars=yes,menubar=no,location=no,resizable=no'
theURL2 = theURL+'?update='+type
Win = window.open(theURL2, "CtrlWindow2", winProp)
if (parseInt(navigator.appVersion) >= 4) { Win.window.focus(); }

}
</script>
<script language="javascript">
function setCuenta(AcctCode, AcctName, Update)
{
	if (Update == "cash")
	{
		document.Form1.CashAcct.value = AcctCode;
		document.Form1.CashAcctName.value = AcctName;
	}
	else if (Update == "check")
	{
		document.Form1.CheckAcct.value = AcctCode;
		document.Form1.CheckAcctName.value = AcctName;
	}
}

function valFrm()
{
	<% If obj <> 2 and obj <> 4 and obj <> 33 Then %>
	if (document.Form1.Series.selectedIndex == 0 && document.Form1.Active.checked)
	{
		alert("<%=getadminDocConfLngStr("LtxtValSelSeries")%>");
		document.Form1.Series.focus();
		return false;
	}
	<% End If %>
	<% If obj = 48 Then %>
	if (document.Form1.Series2.selectedIndex == 0 && document.Form1.Active.checked)
	{
		alert("<%=getadminDocConfLngStr("LtxtValSelSeriesRCT")%>");
		document.Form1.Series2.focus();
		return false;
	}
	<% End If %>
	<% If obj = 24 or obj = 48 Then %>
	if (document.Form1.valChecksFilter.value == 'Y' && document.Form1.ChecksFilter.value != '')
	{
		alert('<%=getadminDocConfLngStr("LtxtValChkFltQryVal")%>');
		document.Form1.btnVerfyFilter.focus();
		return false;
	}
	<% End If %>
	<% If obj = 2 or obj = 4 Then %>
	if (document.Form1.ChkAutGen.checked && document.Form1.AutoGenQry.value == '')
	{
		alert('<%=getadminDocConfLngStr("LtxtValAutoGenQry")%>');
		document.Form1.AutoGenQry.focus();
		return false;
	}
	if (document.Form1.valAutoFilter.value == 'Y' && document.Form1.AutoGenQry.value != '')
	{
		alert('<%=getadminDocConfLngStr("LtxtAutoGenQryVal")%>');
		document.Form1.btnVerfyAutoFilter.focus();
		return false;
	}
	<% End If %>
	return true;
}
</script>
<% If Session("style") = "nc" Then %>
<br>
<% End If %>
<form method="POST" action="adminDocConf.asp" name="Form1" onsubmit="javascript:return valFrm()">
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminDocConfLngStr("LttlConfObjs")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">
		<font face="Verdana" size="1" color="#4783C5"><%=getadminDocConfLngStr("LttlConfObjsNote")%></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td class="style1" colspan="2">
				<select size="1" name="ObjectCode" onchange="javascript:window.location.href='adminDocConf.asp?object='+this.value" class="input">
				<option value=""><%=getadminDocConfLngStr("LtxtSelObj")%></option>
				<optgroup label="<%=getadminDocConfLngStr("LtxtGeneral")%>">
					<option <% If obj = 33 Then %>selected<% End If %> value="33">
					<%=getadminDocConfLngStr("DtxtActivity")%></option>
					<option <% If obj = 4 Then %>selected<% End If %> value="4">
					<%=getadminDocConfLngStr("DtxtItem")%></option>
					<option <% If obj = 2 Then %>selected<% End If %> value="2">
					<%=getadminDocConfLngStr("DtxtBPS")%></option>	
					<option <% If obj = 97 Then %>selected<% End If %> value="97">
					<%=getadminDocConfLngStr("DtxtSO")%></option>					
				</optgroup>
				<optgroup label="<%=getadminDocConfLngStr("LtxtPurchase")%>">
					<option <% If obj = 540000006 Then %>selected<% End If %> value="540000006">
					<%=getadminDocConfLngStr("DtxtPurQuote")%></option>
					<option <% If obj = 20 Then %>selected<% End If %> value="20">
					<%=getadminDocConfLngStr("LtxtGoodsIssue")%></option>
					<option <% If obj = 22 Then %>selected<% End If %> value="22">
					<%=getadminDocConfLngStr("DtxtPurOrder")%></option>
					<% If 1 = 2 Then %><option <% If obj = -22 Then %>selected<% End If %> value="-22">
					<%=getadminDocConfLngStr("DtxtRFM")%></option><% End If %>
				</optgroup>
				<optgroup label="<%=getadminDocConfLngStr("LtxtSale")%>">
					<option <% If obj = 23 Then %>selected<% End If %> value="23">
					<%=getadminDocConfLngStr("DtxtQuote")%></option>
					<option <% If obj = 17 Then %>selected<% End If %> value="17">
					<%=getadminDocConfLngStr("DtxtSalesOrder")%></option>
					<option <% If obj = 15 Then %>selected<% End If %> value="15">
					<%=getadminDocConfLngStr("DtxtDelivery")%></option>
					<option <% If obj = 203 Then %>selected<% End If %> value="203">
					<%=getadminDocConfLngStr("DtxtARDownPayReq")%></option>
					<option <% If obj = 204 Then %>selected<% End If %> value="204">
					<%=getadminDocConfLngStr("DtxtARDownPayInv")%></option>
					<option <% If obj = 13 Then %>selected<% End If %> value="13">
					<%=getadminDocConfLngStr("DtxtInvoice")%></option>
					<option <% If obj = -13 Then %>selected<% End If %> value="-13">
					<%=getadminDocConfLngStr("DtxtInvoice")%> (<%=getadminDocConfLngStr("DtxtReservada")%>)</option>
					<option <% If obj = 48 Then %>selected<% End If %> value="48">
					<%=getadminDocConfLngStr("DtxtInvoice")%>/<%=getadminDocConfLngStr("DtxtReceipt")%></option>
				</optgroup>
				<optgroup label="<%=getadminDocConfLngStr("LtxtBanks")%>">
					<option <% If obj = 24 Then %>selected<% End If %> value="24">
					<%=getadminDocConfLngStr("DtxtReceipt")%></option>
				</optgroup>
				<optgroup label="<%=getadminDocConfLngStr("DtxtService")%>">
					<option <% If obj = 191 Then %>selected<% End If %> value="191">
					<%=getadminDocConfLngStr("DtxtServiceCall")%>
					</option>
					<option <% If obj = 190 Then %>selected<% End If %> value="190">
					<%=getadminDocConfLngStr("DtxtServiceContract")%>
					</option>
					<option <% If obj = 176 Then %>selected<% End If %> value="176">
					<%=getadminDocConfLngStr("DtxtEquipmentCard")%>
					</option>
				</optgroup>
				<optgroup label="<%=getadminDocConfLngStr("DtxtInventory")%>">
					<option <% If obj = 1250000001 Then %>selected<% End If %> value="1250000001">
					<%=getadminDocConfLngStr("DtxtInvTransReq")%>
					</option>
					<option <% If obj = 67 Then %>selected<% End If %> value="67">
					<%=getadminDocConfLngStr("DtxtInvTrans")%>
					</option>
				</optgroup>
				</select></td>
			</tr>
			<% If obj <> -1 Then
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetAdminDocConfData" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@ObjID") = obj
			set rs = Server.CreateObject("ADODB.RecordSet")
			rs.open cmd, , 3, 1 %>
			<% If obj <> 20 Then %>
			<tr>
				<td class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> </font>
				<input type="checkbox" name="Active" value="Y" <% If rs("Active") = "Y" Then %>checked<% End If %> id="Active" class="noborder" style="width: 20px"><font color="#4783C5" face="Verdana" size="1"><label for="Active"><%=getadminDocConfLngStr("DtxtActive")%> 
				(<%=getadminDocConfLngStr("DtxtAgents")%>)</label></font></td>
				<td class="style2">
				&nbsp;</td>
			</tr>
			<% If obj = 2 or obj = 4 Then
			Select Case obj
				Case 2
					ChkAutGen = myApp.AutoGenOCRD
					VarDesc = "Card"
				Case 4
					ChkAutGen = myApp.AutoGenOITM
					VarDesc = "Item"
			End Select %>
			<tr>
				<td class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> </font>
				<input type="checkbox" name="ChkAutGen" value="Y" <% If ChkAutGen Then %>checked<% End If %> id="ChkAutGen" class="noborder" style="width: 20px"><font color="#4783C5" face="Verdana" size="1"><label for="ChkAutGen"><%=getadminDocConfLngStr("LtxtAutoGen")%></label></font></td>
				<td class="style2">
				&nbsp;</td>
			</tr>
			<% set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKAdminGetAutoGenQry" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@ObjCode") = obj
				set rd = cmd.execute()
				%>
			<tr>
				<td valign="top" class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font face="Verdana" color="#4783c5" size="1"> </font>
				<font face="Verdana" size="1" color="#4783C5"><%=getadminDocConfLngStr("DtxtQuery")%> - (set @<%=VarDesc%>Code = (Query))</font></td>
				<td class="style2">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td rowspan="2">
							<textarea rows="10" name="AutoGenQry" dir="ltr" cols="87" onkeydown="javascript:document.Form1.btnVerfyAutoFilter.src='images/btnValidate.gif';document.Form1.btnVerfyAutoFilter.style.cursor = 'hand';;document.Form1.valAutoFilter.value='Y';"><% If Not IsNull(rd("GenQry")) Then %><%=Server.HTMLEncode(rd("GenQry"))%><% End If %></textarea>
						</td>
						<td valign="top">
							<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteAutoFilter" alt="|D:txtDefinition|" onclick="javascript:doFldNote(3, 'AutoGenQry', -1, null);">
						</td>
					</tr>
					<tr>
						<td valign="bottom">
							<img src="images/btnValidateDis.gif" id="btnVerfyAutoFilter" alt="<%=getadminDocConfLngStr("DtxtValidate")%>" onclick="javascript:if (document.Form1.valAutoFilter.value == 'Y')VerfyAutoFilter();">
							<input type="hidden" name="valAutoFilter" value="N">
						</td>
					</tr>
				</table>
				</td>
			</tr>
				<tr>
					<td valign="top" class="style1" style="width: 300px">
					<img src="images/ganchito.gif"><font face="Verdana" color="#4783c5" size="1"> </font><font size="1" color="#4783C5" face="Verdana"><%=getadminDocConfLngStr("LtxtAvlVars")%></font></td>
					<td class="style2">
					<font size="1" color="#4783C5" face="Verdana">
					<span dir="ltr">@LogNum</span> = <%=getadminDocConfLngStr("DtxtLogNum")%></font></td>
				</tr>
			<% End If %>
			<% If obj = 15 or obj = 17 or obj = 23 or obj = 48 Then %>
			<tr>
				<td class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> </font>
				<input type="checkbox" name="ActiveClient" value="Y" <% If rs("ActiveClient") = "Y" Then %>checked<% End If %> id="ActiveClient" class="noborder" style="width: 20px"><font color="#4783C5" face="Verdana" size="1"><label for="ActiveClient"><%=getadminDocConfLngStr("DtxtActive")%> 
				(<%=getadminDocConfLngStr("DtxtClients")%>)</label></font></td>
				<td class="style2">
				&nbsp;</td>
			</tr>
			<% If obj = 48 Then %>
			<tr>
				<td class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> </font>
				<input type="checkbox" name="ClientReservedInvoice" value="Y" <% If myApp.ClientReservedInvoice Then %>checked<% End If %> id="ClientReservedInvoice" class="noborder" style="width: 20px"><font color="#4783C5" face="Verdana" size="1"><label for="ClientReservedInvoice"><%=getadminDocConfLngStr("LtxtClientReservedInv")%> 
				(<%=getadminDocConfLngStr("DtxtClients")%>)</label></font></td>
				<td class="style2">
				&nbsp;</td>
			</tr>
			<% End If %>
			<% End If %>
			<% End If %>
			<% If obj <> 33 and obj <> 20 and obj <> 97 Then %>
			<input type="hidden" name="Confirm" value="<%=rs("Confirm")%>">
			<% End If %>
			<% If obj <> 2 and obj <> 4 and obj <> 33 and obj <> 97 and obj <> -22 Then %>
			<tr>
				<td class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> </font>
				<input type="checkbox" name="ViewObjPrint" value="Y" <% If rs("ViewObjPrint") = "Y" Then %>checked<% End If %> id="ViewObjPrint" class="noborder" style="width: 20px"><font color="#4783C5" face="Verdana" size="1"><label for="ViewObjPrint"><%=getadminDocConfLngStr("LtxtViewObjBtn")%> 
				(<%=getadminDocConfLngStr("DtxtAgents")%>)</label></font></td>
				<td class="style2">
				&nbsp;</td>
			</tr>
			<% If obj <> 203 and obj <> 204 Then %>
			<tr>
				<td class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> </font>
				<input type="checkbox" name="ViewObjPrintClient" value="Y" <% If rs("ViewObjPrintClient") = "Y" Then %>checked<% End If %> id="ViewObjPrintClient" class="noborder" style="width: 20px"><font color="#4783C5" face="Verdana" size="1"><label for="ViewObjPrintClient"><%=getadminDocConfLngStr("LtxtViewObjBtn")%> 
				(<%=getadminDocConfLngStr("DtxtClients")%>)</label></font></td>
				<td class="style2">
				&nbsp;</td>
			</tr>
			<% End If
			If obj <> 176 and obj <> 190 Then %>
			<tr>
				<td class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> 
				<font color="#4783C5"><% If obj <> 48 Then %><%=getadminDocConfLngStr("DtxtSeries")%><% Else %><%=getadminDocConfLngStr("LtxtInvSeries")%><% End If %> (<%=getadminDocConfLngStr("DtxtAgents")%>)</font></font></td>
				<td class="style2">
				<select size="1" name="Series" class="input">
				<option value=""><%=getadminDocConfLngStr("LtxtSelSeries")%></option>
					<%  
					Select Case obj
						Case 48, -13
							Object = 13 
						Case 204
							Object = 203
						Case Else 
							Object = obj
					End Select
					GetQuery rd, 4, Object, null
						do While NOT RD.EOF %>
						<option <% If rd("Series") = rs("Series") Then %>selected<%end if%> value="<%=rd("Series")%>"><%=myHTMLEncode(rd("SeriesName"))%></option>
					<%  RD.MoveNext
					loop    %>
				</select></td>
			</tr>
			<% End If
			If obj = 15 or obj = 17 or obj = 23 or obj = 48 Then %>
			<tr>
				<td class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> 
				<font color="#4783C5"><% If obj <> 48 Then %><%=getadminDocConfLngStr("DtxtSeries")%><% Else %><%=getadminDocConfLngStr("LtxtInvSeries")%><% End If %> (<%=getadminDocConfLngStr("DtxtClients")%>)</font></font></td>
				<td class="style2">
				<select size="1" name="SeriesClient" class="input">
				<option value=""><%=getadminDocConfLngStr("LtxtSelSeries")%></option>
					<%  If obj = 48 Then Object = 13 Else Object = obj
					GetQuery rd, 4, Object, null
						do While NOT RD.EOF %>
						<option <% If rd("Series") = rs("SeriesClient") Then %>selected<%end if%> value="<%=rd("Series")%>"><%=myHTMLEncode(rd("SeriesName"))%></option>
					<%  RD.MoveNext
					loop    %>
				</select></td>
			</tr>
			<% End If %>
			<% End If %>
			<% If obj = 48 Then %>
			<tr>
				<td class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> 
				<font color="#4783C5"><%=getadminDocConfLngStr("LtxtRctSeries")%> (<%=getadminDocConfLngStr("DtxtAgents")%>)</font></font></td>
				<td class="style2">
				<select size="1" name="Series2" class="input">
				<option value=""><%=getadminDocConfLngStr("LtxtSelSeries")%></option>
					<%  
					If Not IsNull(rs("Series2")) Then Series2 = CInt(rs("Series2")) Else Series2 = -1
					GetQuery rd, 4, 24, null
						do While NOT RD.EOF %>
						<option <% If CInt(rd("Series")) = Series2 Then %>selected<%end if%> value="<%=rd("Series")%>"><%=myHTMLEncode(rd("SeriesName"))%></option>
					<%  RD.MoveNext
					loop    %>
				</select></td>
			</tr>
			<% If obj <> 13 Then %>
			<tr>
				<td class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> 
				<font color="#4783C5"><%=getadminDocConfLngStr("LtxtRctSeries")%> (<%=getadminDocConfLngStr("DtxtClients")%>)</font></font></td>
				<td class="style2">
				<select size="1" name="Series2Client" class="input">
				<option value=""><%=getadminDocConfLngStr("LtxtSelSeries")%></option>
					<%  
					If Not IsNull(rs("Series2Client")) Then Series2Client = CInt(rs("Series2Client")) Else Series2 = -1
					GetQuery rd, 4, 24, null
						do While NOT RD.EOF %>
						<option <% If CInt(rd("Series")) = Series2Client Then %>selected<%end if%> value="<%=rd("Series")%>"><%=myHTMLEncode(rd("SeriesName"))%></option>
					<%  RD.MoveNext
					loop    %>
				</select></td>
			</tr>
			<% End If %>
			<% End If %>
			<% If obj = 17 Then %>
			<% If myApp.Enable3dx Then %>
			<tr>
				<td class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font color="#4783C5" face="Verdana" size="1">
				<%=getadminDocConfLngStr("LtxtUn3dxVer")%></font></td>
				<td class="style2">
				<select size="1" name="Verfy3dx" class="input">
				<option value="N"><%=getadminDocConfLngStr("DtxtDisabled")%></option>
				<option <% If myApp.Verfy3dxOrder = "C" Then %>selected<% End If %> value="C">
				<%=getadminDocConfLngStr("DtxtConfirm")%></option>
				<option <% If myApp.Verfy3dxOrder = "O" Then %>selected<% End If %> value="O">
				<%=getadminDocConfLngStr("DtxtNotNull")%></option>
				</select></td>
			</tr>
			<% End If %>
			<tr>
				<td class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font color="#4783C5" face="Verdana" size="1">
				<%=getadminDocConfLngStr("LtxtInBtchVer")%></font></td>
				<td class="style2">
				<select size="1" name="VerfyBtch" class="input">
				<option value="N"><%=getadminDocConfLngStr("DtxtDisabled")%></option>
				<option <% If myApp.VerfyBtchOrder = "C" Then %>selected<% End If %> value="C">
				<%=getadminDocConfLngStr("DtxtConfirm")%></option>
				<option <% If myApp.VerfyBtchOrder = "O" Then %>selected<% End If %> value="O">
				<%=getadminDocConfLngStr("DtxtNotNull")%></option>
				</select></td>
			</tr>
			<% End If %>
			<% If obj <> 2 and obj <> 4 and obj <> 33 and obj <> 20 and obj <> 97 and obj <> -22 and obj <> 67 and obj <> 176 and obj <> 190 and obj <> 191 and obj <> 1250000001 Then %>
			<tr>
				<td class="style1" style="width: 300px; height: 22px;">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> </font>
				<input type="checkbox" <% If rs("PrintCmpPaper") = "Y" Then %>checked<% End If %> name="PrintCmpPaper" value="Y" id="PrintCmpPaper" class="noborder"><font color="#4783C5" face="Verdana" size="1"><label for="PrintCmpPaper"><%=getadminDocConfLngStr("LtxtPrintCmpPaper")%></label></font></td>
				<td class="style2" style="height: 22px"></td>
			</tr>
			<tr>
				<td class="style1" style="width: 300px; padding-top: 2px;" valign="top">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> </font>
				<font color="#4783C5" face="Verdana" size="1"><%=getadminDocConfLngStr("LtxtObjNote")%></font></td>
				<td class="style2">
					<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td><%
									Dim oFCKeditor
									Set oFCKeditor = New FCKeditor
									oFCKeditor.BasePath = "FCKeditor/"
									oFCKeditor.Height = 300
									oFCKEditor.ToolbarSet = "Custom"
									If Not IsNull(rs("Note")) Then oFCKEditor.Value = myHTMLEncode(rs("Note"))
									oFCKEditor.Config("AutoDetectLanguage") = False
									If Session("myLng") <> "pt" Then
										oFCKEditor.Config("DefaultLanguage") = Session("myLng")
									Else
										oFCKEditor.Config("DefaultLanguage") = "pt-br"
									End If
									oFCKeditor.Create "txtNote"
									%>
						</td>
						<td width="16" valign="bottom">
						<a href="javascript:doFldTrad('DocConf', 'ObjectCode', '<%=obj%>', 'AlterNote', 'R', null);"><img src="images/trad.gif" alt="<%=getadminDocConfLngStr("DtxtTranslate")%>" border="0"></a>
						</td>
					</tr>
					</table>
				</td>
			</tr>
			<% End If %>
			<% If obj = 24 or obj = 48 Then
			If obj = 24 Then %>
			<tr>
				<td class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> </font>
				<input type="checkbox" <% If rs("PrintContact") = "Y" Then %>checked<% End If %> name="PrintContact" value="Y" id="PrintContact" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="PrintContact"><%=getadminDocConfLngStr("LtxtPrintContact")%></label></font></td>
				<td class="style2">&nbsp;</td>
			</tr>
			<tr>
				<td class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> </font>
				<input type="checkbox" <% If rs("PrintAddress") = "Y" Then %>checked<% End If %> name="PrintAddress" value="Y" id="PrintAddress" class="noborder"><font color="#4783C5" face="Verdana" size="1"><label for="PrintAddress"><%=getadminDocConfLngStr("LtxtPrintAddress")%></label></font></td>
				<td class="style2">&nbsp;</td>
			</tr>
			<% End If %>
			<tr>
				<td class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> </font>
				<input type="checkbox" <% If rs("PrintPaidSum") = "Y" Then %>checked<% End If %> name="PrintPaidSum" value="Y" id="PrintPaidSum" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="PrintPaidSum"><%=getadminDocConfLngStr("LtxtPrintPaidSum")%></label></font></td>
				<td class="style2">&nbsp;</td>
			</tr>
			<% 
			If Not myApp.SegAct Then
				DispCash = "CashAcct"
				DispCheck = "CheckAcct"
			Else
				DispCash = "OLKCommon.dbo.DBOLKGetSegmentAccount" & Session("ID") & "(CashAcct)"
				DispCheck = "OLKCommon.dbo.DBOLKGetSegmentAccount" & Session("ID") & "(CheckAcct)"
			End If
			sql = 	"select ChecksFilter, " & _
					"CashAcct, " & DispCash & " DispCashAcct, CheckAcct, " & DispCheck & " DispCheckAcct, IsNull((select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OACT', 'AcctName', T1.CashAcct, AcctName) from OACT where AcctCode = T1.CashAcct), '') CashAcctName, " & _
					"IsNull((select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OACT', 'AcctName', T1.CheckAcct, AcctName) from OACT where AcctCode = T1.CheckAcct), '') CheckAcctName " & _
					"from olkcommon T0 " & _
					"cross join OLKDocConf T1 " & _
					"where ObjectCode = " & obj
			set rs = conn.execute(sql) %>
			<tr>
				<td class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> </font>
				<input type="checkbox" <% If myApp.ORCTContraComp Then %>checked<% End If %> name="ORCTContraComp" value="Y" id="ORCTContraComp" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="ORCTContraComp"><%=getadminDocConfLngStr("LtxtORCTContraComp")%></label></font></td>
				<td class="style2">&nbsp;</td>
			</tr>
			<tr>
				<td class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> </font>
				<input type="checkbox" <% If myApp.ApplyOpenRctToInvBal Then %>checked<% End If %> name="ApplyOpenRctToInvBal" value="Y" id="ApplyOpenRctToInvBal" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="ApplyOpenRctToInvBal"><%=getadminDocConfLngStr("LtxtApplyOpenRctToInv")%></label></font></td>
				<td class="style2">&nbsp;</td>
			</tr>
			<tr>
				<td class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font face="Verdana" color="#4783c5" size="1"> </font>
				<font face="Verdana" size="1" color="#4783C5"><%=getadminDocConfLngStr("LtxtDefCashAcct")%></font></td>
				<td class="style2">
				<% 
				Dim myAccount
				set myAccount = New AccountControl
				myAccount.ID = "CashAcct"
				myAccount.Value = rs("CashAcct")
				myAccount.DisplayValue = rs("DispCashAcct")
				myAccount.Description = rs("CashAcctName")
				myAccount.AccountType = "cash"
				myAccount.GenerateAccount %>
				</td>
			</tr>
			<tr>
				<td class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font face="Verdana" color="#4783c5" size="1"> </font>
				<font face="Verdana" size="1" color="#4783C5"><%=getadminDocConfLngStr("LtxtDefChkAcct")%></font></td>
				<td class="style2">
				<% 
				set myAccount = New AccountControl
				myAccount.ID = "CheckAcct"
				myAccount.Value = rs("CheckAcct")
				myAccount.DisplayValue = rs("DispCheckAcct")
				myAccount.Description = rs("CheckAcctName")
				myAccount.AccountType = "check"
				myAccount.GenerateAccount %>
				</td>
			</tr>
			<tr>
				<td class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> </font>
				<input type="checkbox" <% If myApp.IgnoreSystemChecksFilter Then %>checked<% End If %> name="IgnoreSystemChecksFilter" value="Y" id="IgnoreSystemChecksFilter" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="IgnoreSystemChecksFilter"><%=getadminDocConfLngStr("LtxtIgnoreSystemCheck")%></label></font></td>
				<td class="style2">&nbsp;</td>
			</tr>
			<tr>
				<td valign="top" class="style1" style="width: 300px">
				<img src="images/ganchito.gif"><font face="Verdana" color="#4783c5" size="1"> </font>
				<font face="Verdana" size="1" color="#4783C5"><%=getadminDocConfLngStr("LtxtChecksFilterQry")%> - (AcctCode not in)</font></td>
				<td class="style2">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td rowspan="2">
							<textarea rows="10" name="ChecksFilter" dir="ltr" cols="87" onkeydown="javascript:document.Form1.btnVerfyFilter.src='images/btnValidate.gif';document.Form1.btnVerfyFilter.style.cursor = 'hand';;document.Form1.valChecksFilter.value='Y';"><% If Not IsNull(myApp.ChecksFilter) Then %><%=Server.HTMLEncode(myApp.ChecksFilter)%><% End If %></textarea>
						</td>
						<td valign="top">
							<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="|D:txtDefinition|" onclick="javascript:doFldNote(3, 'ChecksFilter', -1, null);">
						</td>
					</tr>
					<tr>
						<td valign="bottom">
							<img src="images/btnValidateDis.gif" id="btnVerfyFilter" alt="<%=getadminDocConfLngStr("DtxtValidate")%>" onclick="javascript:if (document.Form1.valChecksFilter.value == 'Y')VerfyFilter();">
							<input type="hidden" name="valChecksFilter" value="N">
						</td>
					</tr>
				</table>
				</td>
			</tr>
				<tr>
					<td valign="top" class="style1" style="width: 300px">
					<img src="images/ganchito.gif"><font face="Verdana" color="#4783c5" size="1"> </font><font size="1" color="#4783C5" face="Verdana"><%=getadminDocConfLngStr("LtxtAvlVars")%></font></td>
					<td class="style2">
					<font size="1" color="#4783C5" face="Verdana">
					<span dir="ltr">@branch</span> = <%=getadminDocConfLngStr("DtxtBranch")%><br>
					<span dir="ltr">@SlpCode</span> = <%=getadminDocConfLngStr("LtxtAgentCode")%></font></td>
				</tr>
			<% End If %>
			<% End If %>
			</table>
		</td>
	</tr>
	<% If obj <> -1 Then %>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminDocConfLngStr("DtxtSave")%>" name="B1" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<% End If %>
	</table>
<input type="hidden" name="object" value="<%=obj%>">
<input type="hidden" name="save" value="True">
</form>
<script language="javascript" src="accountControl.js"></script>
<script language="javascript">
var verfyButton;
var hdverfyButton;
function VerfyFilter()
{
	verfyButton = document.Form1.btnVerfyFilter;
	hdverfyButton = document.Form1.valChecksFilter;
	document.frmVerfyQuery.type.value = 'ChecksFilter';
	document.frmVerfyQuery.Query.value = document.Form1.ChecksFilter.value;
	if (document.frmVerfyQuery.Query.value != '')
	{
		document.frmVerfyQuery.submit();
	}
	else
	{
		VerfyQueryVerified();
	}
}
function VerfyAutoFilter()
{
	verfyButton = document.Form1.btnVerfyAutoFilter;
	hdverfyButton = document.Form1.valAutoFilter;
	document.frmVerfyQuery.type.value = 'AutoGenCode';
	document.frmVerfyQuery.Query.value = document.Form1.AutoGenQry.value;
	if (document.frmVerfyQuery.Query.value != '')
	{
		document.frmVerfyQuery.submit();
	}
	else
	{
		VerfyQueryVerified();
	}
}
function VerfyQueryVerified()
{
	verfyButton.src='images/btnValidateDis.gif'
	verfyButton.style.cursor = '';
	hdverfyButton.value='N';
}
//-->
</script>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
	<input type="hidden" name="type" value="">
	<input type="hidden" name="obj" value="<%=obj%>">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="parent" value="Y">
</form>
<!--#include file="bottom.asp" -->