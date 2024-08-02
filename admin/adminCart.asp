<!--#include file="top.asp" -->
<!-- #INCLUDE file="FCKeditor/fckeditor.asp" -->
<!--#include file="lang/adminCart.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<head>
<% 
set rd = Server.CreateObject("ADODB.RecordSet")
sql = "select (select CardName from OCRD where CardCode = T0.AnonCartClient) from OLKCommon T0 "
set rd = conn.execute(sql)
AnonCartClientName = rd(0)
rd.close
sql = "select ObjectCode, ActiveClient from OLKDocConf where ObjectCode in (15, 17, 23, 48)"
rd.open sql, conn, 3, 1 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<SCRIPT LANGUAGE="JavaScript">
function noteEdit(cmd) 
{
	var noteIndex = ''
	if (document.frmAdminCart.Notes.selectedIndex > 0)
	{
		noteIndex = document.frmAdminCart.Notes.value
	}
	if (noteIndex == '' & cmd == "e") { } 
	else 
	{
		var page = 'adminNote.asp?noteIndex=' + noteIndex + '&cmd=' + cmd + '&pop=Y'
		OpenWin = this.open(page, "NoteEdit", "toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no, width=360,height=140");
	}
}

function Start2(theURL, popW, popH, type) { // V 1.0
var winleft = (screen.width - popW) / 2;
var winUp = (screen.height - popH) / 2;
winProp = 'width='+popW+',height='+popH+',left='+winleft+',top='+winUp+',toolbar=no,scrollbars=yes,menubar=no,location=no,resizable=no'
theURL2 = theURL+'?update='+type
Win = window.open(theURL2, "CtrlWindow2", winProp)
if (parseInt(navigator.appVersion) >= 4) { Win.window.focus(); }

}
function chkMaxDisc(Field, OldVal)
{
	if (!IsNumeric(Field.value))
	{
		alert("<%=getadminCartLngStr("DtxtValNumVal")%>");
		Field.value = OldVal;
	}
	else if (parseFloat(Field.value) > 100)
	{
		alert("<%=getadminCartLngStr("DtxtValNumMaxVal")%>".replace("{0}", "100"));
		Field.value = 100;
	}
	else if (parseFloat(Field.value) < 1)
	{
		alert("<%=getadminCartLngStr("DtxtValNumMinVal")%>".replace("{0}", "1"));
		Field.value = 1;
	}
	Field.value = formatNumber(Field.value, <%=myApp.PercentDec%>);
}
function delNote()
{
	if (document.frmAdminCart.Notes.selectedIndex > 0)
	{
		if (confirm('<%=getadminCartLngStr("LtxtConfDelNote")%>'))
		{
			window.location.href = 'adminSubmit.asp?submitCmd=adminnote&cmd=d&redir=general&noteIndex=' + document.frmAdminCart.Notes.value;
		}
	}
}
function valFrm()
{
	if (document.frmAdminCart.EnableAnonCart.checked && document.frmAdminCart.AnonCartClient.value == '')
	{
		alert('<%=getadminCartLngStr("LtxtSelClient")%>');
		return false;
	}
	return true;
}
</script>
<script language="javascript" src="js_up_down.js"></script>
</head>

<form method="POST" action="adminsubmit.asp" name="frmAdminCart" onsubmit="javascript:return valFrm();">
<%
strFormName = "frmAdminCart"
strTextAreaName = "CCartNote"
%>
<table border="0" cellpadding="0" width="100%" id="table3">
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminCartLngStr("LttlCartProp")%> </font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
		</font>
		<font face="Verdana" size="1" color="#4783C5"><%=getadminCartLngStr("LttlCartPropNote")%></font></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE" height="269">
		<div align="center">
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font><font face="Verdana" size="1" color="#4783C5">
					<%=getadminCartLngStr("LtxtCartType")%></font></td>
					<td>
					<p>
					<select size="1" name="CartType" class="input" onchange="changeCartType(this.value);">
					<option value="A" <% If myApp.CartType = "A" Then %>selected<%end if%>>
					<%=getadminCartLngStr("LtxtApplication")%></option>
					<option value="S" <% If myApp.CartType = "S" Then %>selected<%end if%>>
					<%=getadminCartLngStr("LtxtSite")%></option>
					</select></td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="300">
					<img src="images/ganchito.gif"><font color="#4783C5"> </font>
					<input type="checkbox" class="noborder" name="EnableAnonCart" value="Y" <% If myApp.EnableAnonCart Then %>checked<%end if %> id="EnableAnonCart"><font face="Verdana" size="1" color="#4783C5"><label for="EnableAnonCart"><%=getadminCartLngStr("LtxtEnableAnonCart")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="300">
					<img src="images/ganchito.gif"><font color="#4783C5">
					</font><font face="Verdana" size="1" color="#4783C5">
					<%=getadminCartLngStr("LtxtAnonCartClient")%></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					<input type="text" name="AnonCartClient" size="15" class="input" dir="ltr" value="<%=myApp.AnonCartClient%>" onkeydown="return chkMax(event, this, 15);"><input type="text" readonly name="AnonCartClientName" size="50" class="inputDis" value="<%=myHTMLEncode(AnonCartClientName)%>"></td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif"> 
					<input <% If myApp.EnableHideCartHdr Then %> checked <% End If %> type="checkbox" name="EnableHideCartHdr" value="Y" id="EnableHideCartHdr" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnableHideCartHdr"><%=getadminCartLngStr("LtxtEnableHideCartHdr")%></label></font></td>
					<td>
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font><font face="Verdana" size="1" color="#4783C5">
					<%=getadminCartLngStr("LtxtAfterCartAddC")%></font></td>
					<td>
					<p>
					<select size="1" name="AfterCartAddC" class="input">
					<option value="Y" <% If myApp.AfterCartAddC = "Y" Then %>selected<%end if%>>
					<%=getadminCartLngStr("LtxtOptGoToCart")%></option>
					<option value="N" <% If myApp.AfterCartAddC = "N" Then %>selected<%end if%>>
					<%=getadminCartLngStr("LtxtOptCurPag")%></option>
					</select></td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					<font color="#4783C5"><%=getadminCartLngStr("LtxtAfterCartAddV")%></font></font></td>
					<td>
					<p>
					<select size="1" name="AfterCartAddV" class="input">
					<option value="Y" <% If myApp.AfterCartAddV = "Y" Then %>selected<%end if%>>
					<%=getadminCartLngStr("LtxtOptGoToCart")%></option>
					<option value="N" <% If myApp.AfterCartAddV = "N" Then %>selected<%end if%>>
					<%=getadminCartLngStr("LtxtOptCurPag")%></option>
					</select></td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					<font color="#4783C5"><%=getadminCartLngStr("LtxtAfterCartAddPocke")%></font></font></td>
					<td>
					<p>
					<select size="1" name="AfterCartAddPocket" class="input">
					<option value="N" <% If myApp.AfterCartAddPocket = "N" Then %>selected<%end if%>>
					<%=getadminCartLngStr("LtxtOptGoToCart")%></option>
					<option value="Y" <% If myApp.AfterCartAddPocket = "Y" Then %>selected<%end if%>>
					<%=getadminCartLngStr("LtxtBackToSearch")%></option>
					</select></td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif">
					<input type="checkbox" <% If myApp.CartItmBarCode Then %>checked<% End If %> name="CartItmBarCode" value="Y" id="CartItmBarCode" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="CartItmBarCode"><%=getadminCartLngStr("LtxtCartItmBarCode")%></label></font></td>
					<td>
					&nbsp;</td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif">
					<input type="checkbox" <% If myApp.FastAddUnRem Then %>checked<% End If %> name="FastAddUnRem" value="Y" id="FastAddUnRem" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="FastAddUnRem"><%=getadminCartLngStr("LtxtFastAddUnRem")%></label></font></td>
					<td>
					&nbsp;</td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif">
					<input type="checkbox" <% If myApp.FastAddBeep Then %>checked<% End If %> name="FastAddBeep" value="Y" id="FastAddBeep" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="FastAddBeep"><%=getadminCartLngStr("LtxtFastAddBeep")%></label></font></td>
					<td>
					&nbsp;</td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif"><font face="Verdana" size="1" color="#4783C5"> 
					<%=getadminCartLngStr("LtxtDefCDoc")%></font></td>
					<td>
					<p>
					<select size="1" name="D_DocC" class="input">
					<%
					noActiveDoc = False
					rd.Filter = "ObjectCode = 23"
					If rd("ActiveClient") = "Y" Then
					noActiveDoc = True %>
				    <option value="C" <% If myApp.D_DocC = "C" Then Response.Write "selected" %>>
					<%=getadminCartLngStr("DtxtQuote")%></option><% End If
					rd.Filter = "ObjectCode = 17"
					If rd("ActiveClient") = "Y" Then
					noActiveDoc = True %>
					<option value="P" <% If myApp.D_DocC = "P" Then Response.Write "selected" %>>
					<%=getadminCartLngStr("DtxtSalesOrder")%></option><% End If
					rd.Filter = "ObjectCode = 15"
					If rd("ActiveClient") = "Y" Then
					noActiveDoc = True %>
					<option value="E" <% If myApp.D_DocC = "E" Then Response.Write "selected" %>>
					<%=getadminCartLngStr("DtxtDelivery")%></option><% End If
					rd.Filter = "ObjectCode = 48"
					If rd("ActiveClient") = "Y" Then
					noActiveDoc = True
					Active48 = True
					If myApp.CartType = "S" Then %>
					<option value="F" <% If myApp.D_DocC = "F" Then Response.Write "selected" %>>
					<%=getadminCartLngStr("DtxtInvoice")%>/<%=getadminCartLngStr("DtxtReceipt")%></option><% End If %><% End If %>
					<% If Not noActiveDoc Then %>
					<option value="N">
					<%=getadminCartLngStr("DtxtUndefined")%></option><% End If %>
				    </select></td>
				</tr>
				<% If 1 = 2 Then %>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif"><font face="Verdana" size="1" color="#4783C5"> 
					<%=getadminCartLngStr("LtxtDefPDoc")%></font></td>
					<td>
					<p>
					<select size="1" name="PocketDefDoc" class="input">
					<option value="23" <% If myApp.PocketDefDoc = 23 Then Response.Write "selected" %>>
					<%=getadminCartLngStr("DtxtQuote")%></option>
				    <option value="17" <% If myApp.PocketDefDoc = 17 Then Response.Write "selected" %>>
					<%=getadminCartLngStr("DtxtSalesOrder")%></option>
				    </select></td>
				</tr>
				<% Else %>
				<input type="hidden" name="PocketDefDoc" value="<%=myApp.PocketDefDoc%>">
				<% End If %>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif">
					<input type="checkbox" <% If myApp.EnCSelDoc Then %>checked<% End If %> name="EnCSelDoc" value="Y" id="EnCSelDoc" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnCSelDoc"><%=getadminCartLngStr("LtxtEnCSelDoc")%></label></font></td>
					<td>
					&nbsp;</td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif">
					<input type="checkbox" <% If myApp.EnableDocPrjSel Then %>checked<% End If %> name="EnableDocPrjSel" value="Y" id="EnableDocPrjSel" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnableDocPrjSel"><%=getadminCartLngStr("LtxtEnableDocPrjSel")%></label></font></td>
					<td>
					&nbsp;</td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif">
					<input type="checkbox" <% If myApp.AllowClientPartSuppSel Then %>checked<% End If %> name="AllowClientPartSuppSel" value="Y" id="AllowClientPartSuppSel" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="AllowClientPartSuppSel"><%=getadminCartLngStr("LtxtAllowClParSupp")%></label></font></td>
					<td>
					&nbsp;</td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif"> 
					<input <% If myApp.EnableMultiBPCart Then %>checked<% End If %> type="checkbox" name="EnableMultiBPCart" value="Y" id="EnableMultiBPCart" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnableMultiBPCart"><%=getadminCartLngStr("LtxtEnableMultiBPCart")%></label></font></td>
					<td>
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif"> 
					<input <% If myApp.EnableCartSum Then %>checked<% End If %> type="checkbox" name="EnableCartSum" value="Y" id="EnableCartSum" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnableCartSum"><%=getadminCartLngStr("LtxtEnableCartSum")%></label></font></td>
					<td>
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif"> 
					<font face="Verdana" size="1" color="#4783C5">
					<%=getadminCartLngStr("LtxtCartSumQty")%></font></td>
					<td>
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td><input type="text" name="CartSumQty" id="CartSumQty" size="5" class="input" value="<%=myApp.CartSumQty%>" onfocus="this.select()" onkeydown="return chkMax(event, this, 6);"></td>
							<td valign="middle">
							<table cellpadding="0" cellspacing="0" border="0">
								<tr>
									<td><img src="images/img_nud_up.gif" id="btnCartSumQtyUp"></td>
								</tr>
								<tr>
									<td><img src="images/spacer.gif"></td>
								</tr>
								<tr>
									<td><img src="images/img_nud_down.gif" id="btnCartSumQtyDown"></td>
								</tr>
							</table>
							</td>
						</tr>
					</table>
					<script language="javascript">NumUDAttach('frmAdminCart', 'CartSumQty', 'btnCartSumQtyUp', 'btnCartSumQtyDown');</script></td>
				</tr>
				<input type="hidden" name="EnableDiscount" value="<%=myApp.EnableDiscount%>">
				<tr>
					<td width="435">
					<img src="images/ganchito.gif"> 
					<input <% If myApp.ShowPriceBefDiscount Then %>checked<% End If %> type="checkbox" name="ShowPriceBefDiscount" value="Y" id="ShowPriceBefDiscount" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="ShowPriceBefDiscount"><%=getadminCartLngStr("LtxtShowPriceBefDisco")%></label></font></td>
					<td>
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif"> 
					<input <% If myApp.PrintPriceBefDiscount Then %>checked<% End If %> type="checkbox" name="PrintPriceBefDiscount" value="Y" id="PrintPriceBefDiscount" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="PrintPriceBefDiscount"><%=getadminCartLngStr("LtxtPrintPriceBefDisc")%></label></font></td>
					<td>
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif"> 
					<input <% If myApp.ShowLineDiscount Then %>checked<% End If %> type="checkbox" name="ShowLineDiscount" value="Y" id="ShowLineDiscount" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="ShowLineDiscount"><%=getadminCartLngStr("LtxtShowLineDisco")%></label></font></td>
					<td>
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif"> 
					<input <% If myApp.PrintLineDiscount Then %>checked<% End If %> type="checkbox" name="PrintLineDiscount" value="Y" id="PrintLineDiscount" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="PrintLineDiscount"><%=getadminCartLngStr("LtxtPrintLineDisco")%></label></font></td>
					<td>
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif"> 
					<font face="Verdana" size="1" color="#4783C5"><%=getadminCartLngStr("LtxtMaxDiscount")%></font></td>
					<td>
					<p>
					<input type="text" name="MaxDiscount" size="10" class="input" style="text-align: right; " value='<%=FormatNumber(CDbl(myApp.MaxDiscount), myApp.PercentDec)%>' onfocus="this.select()" onchange="chkMaxDisc(this, document.frmAdminCart.oldMaxDiscount.value);document.frmAdminCart.oldMaxDiscount.value=this.value;" onkeydown="return chkMax(event, this, 7);">
					<input type="hidden" name="oldMaxDiscount" id="oldMaxDiscount" value="<%=FormatNumber(CDbl(myApp.MaxDiscount), myApp.PercentDec)%>"></td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif"> 
					<input <% If myApp.ApplyMaxDiscToSU Then %>checked<% End If %> type="checkbox" name="ApplyMaxDiscToSU" value="Y" id="ApplyMaxDiscToSU" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="ApplyMaxDiscToSU"><%=getadminCartLngStr("LtxtApplyMaxDiscToSU")%></label></font></td>
					<td>
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif"> 
					<input <% If myApp.EnableClientMDoc Then %>checked<% End If %> type="checkbox" name="EnableClientMDoc" value="Y" id="EnableClientMDoc" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnableClientMDoc"><%=getadminCartLngStr("LtxtEnableClientMDoc")%></label></font></td>
					<td>
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif"> 
					<font face="Verdana" size="1" color="#4783C5">
					<input <% If myApp.EnableCartImpC Then %>checked<% End If %> type="checkbox" name="EnableCartImpC" value="Y" id="EnableCartImpC" class="noborder"><label for="EnableCartImpC"><%=getadminCartLngStr("LtxtEnableCartImpC")%></label></font></td>
					<td>
					&nbsp;</td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif"> 
					<font face="Verdana" size="1" color="#4783C5">
					<input <% If myApp.EnableCartImpV Then %>checked<% End If %> type="checkbox" name="EnableCartImpV" value="Y" id="EnableCartImpV" class="noborder"><label for="EnableCartImpV"><%=getadminCartLngStr("LtxtEnableCartImpV")%></label></font></td>
					<td>
					<input type="hidden" name="EnTop10Items" value="<%=GetYN(myApp.EnTop10Items)%>">
					<input type="hidden" name="Top10Items" value="<%=myApp.Top10Items%>">
					<input type="hidden" name="CartGroup" value="<%=myApp.CartGroup%>"></td>
				</tr>
				<tr>
					<td width="435">
					<img src="images/ganchito.gif">
					<input <% If myApp.ExpItems Then %>checked<% End If %> type="checkbox" name="ExpItems" value="Y" id="ExpItems" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="ExpItems"><%=getadminCartLngStr("LtxtExpItems")%></label></font></td>
					<td>
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td width="435" valign="top">
					<img src="images/ganchito.gif">
					<font face="Verdana" size="1" color="#4783C5">
					<%=getadminCartLngStr("LtxtCCartNote")%><br>
					</font>
					<img src="images/ganchito.gif">
					<font face="Verdana" size="1" color="#4783C5">
					<input type="checkbox" name="PrintCCartNote" <% If myApp.PrintCCartNote Then %>checked<% End If %> value="Y" id="PrintCCartNote" class="noborder"><label for="PrintCCartNote"><%=getadminCartLngStr("LtxtPrintCCartNote")%></label></font></td>
					<td>
					<table cellpadding="0" cellspacing="0" border="0" width="100%">
						<tr>
							<td>
							<%
							Dim oFCKeditor
							Set oFCKeditor = New FCKeditor
							oFCKeditor.BasePath = "FCKeditor/"
							oFCKeditor.Height = 300
							oFCKEditor.ToolbarSet = "Custom"
							If Not IsNull(myApp.CCartNote) Then oFCKEditor.Value = myApp.CCartNote
							oFCKEditor.Config("AutoDetectLanguage") = False
							If Session("myLng") <> "pt" Then
								oFCKEditor.Config("DefaultLanguage") = Session("myLng")
							Else
								oFCKEditor.Config("DefaultLanguage") = "pt-br"
							End If
							oFCKeditor.Create "CCartNote"
							%>
							</td>
							<td width="16" valign="bottom">
							<a href="javascript:doFldTrad('Common', '', '', 'AlterCCartNote', 'R', null);"><img src="images/trad.gif" alt="<%=getadminCartLngStr("DtxtTranslate")%>" border="0"></a>
							</td>
						</tr>
					</table>
					</td>
					</tr>
					<tr>
					<td width="435" valign="top">
					<img src="images/ganchito.gif">
					<font face="Verdana" size="1" color="#4783C5">
					<%=getadminCartLngStr("LtxtTransConfNote")%><br>
					</font>
					<img src="images/ganchito.gif">
					<font face="Verdana" size="1" color="#4783C5">
					<input type="checkbox" name="UseCustomTransMsg" <% If myApp.UseCustomTransMsg Then %>checked<% End If %> value="Y" id="UseCustomTransMsg" class="noborder"><label for="UseCustomTransMsg"><%=getadminCartLngStr("LtxtUseCustomTransMsg")%></label></font></td>
					<td>
					<table cellpadding="0" cellspacing="0" border="0" width="100%">
						<tr>
							<td>
							<%
							'Dim oFCKeditor
							Set oFCKeditor = New FCKeditor
							oFCKeditor.BasePath = "FCKeditor/"
							oFCKeditor.Height = 300
							oFCKEditor.ToolbarSet = "Custom"
							If Not IsNull(myApp.CustomTransMsg) Then oFCKEditor.Value = myApp.CustomTransMsg
							oFCKEditor.Config("AutoDetectLanguage") = False
							If Session("myLng") <> "pt" Then
								oFCKEditor.Config("DefaultLanguage") = Session("myLng")
							Else
								oFCKEditor.Config("DefaultLanguage") = "pt-br"
							End If
							oFCKeditor.Create "CustomTransMsg"
							%>
							</td>
							<td width="16" valign="bottom">
							<a href="javascript:doFldTrad('Common', '', '', 'AlterCustomTransMsg', 'R', null);"><img src="images/trad.gif" alt="<%=getadminCartLngStr("DtxtTranslate")%>" border="0"></a>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td width="435">
					<img src="images/ganchito.gif">
					<input <% If myApp.BasketMItems Then %>checked<% End If %> type="checkbox" name="BasketMItems" value="Y" id="BasketMItems" class="noborder" onclick="document.getElementById('EnSelAll').disabled=!this.checked;"><label for="BasketMItems"><font face="Verdana" size="1" color="#4783C5"><%=getadminCartLngStr("LtxtBasketMItems")%></font></label></td>
				<td valign="top">
				<p>
					&nbsp;</td>
			</tr>
			<tr>
				<td width="435">
					<img src="images/ganchito.gif">
					<input <% If myApp.EnSelAll Then %>checked<% End If %> <% If Not myApp.BasketMItems Then %>disabled<% End If %> type="checkbox" name="EnSelAll" value="Y" id="EnSelAll" class="noborder"><label for="EnSelAll"><font face="Verdana" size="1" color="#4783C5"><%=getadminCartLngStr("LtxtEnSelAll")%></font></label></td>
				<td valign="top">
				<p>
					&nbsp;</td>
			</tr>
			<tr>
				<td width="435">
					<img src="images/ganchito.gif">
					<font face="Verdana" size="1" color="#4783C5"><%=getadminCartLngStr("LtxtSellAllFrom")%></font></td>
				<td valign="top">
				<p><select size="1" name="EnSellAllUnitFrom" class="input">
					<option value="1" <% If myApp.EnSellAllUnitFrom = 1 Then %>selected<%end if%>>
					<%=getadminCartLngStr("DtxtBaseUnit")%></option>
					<option value="2" <% If myApp.EnSellAllUnitFrom = 2 Then %>selected<%end if%>>
					<%=getadminCartLngStr("DtxtSalUnit")%></option>
					<option value="3" <% If myApp.EnSellAllUnitFrom = 3 Then %>selected<%end if%>>
					<%=getadminCartLngStr("DtxtPackUnit")%></option>
					</select></td>
			</tr>
			<% If 1 = 2 Then %>
			<tr>
				<td width="435">
				<img src="images/ganchito.gif"> 
				<font color="#4783C5" face="Verdana" size="1"><%=getadminCartLngStr("LtxtDocMCBal")%></font></td>
				<td valign="top">
				<p>&nbsp;<select size="1" name="DocMCBal" class="input">
				<option <% If myApp.DocMCBal = "L" Then %>selected<% End If %> value="L">
				<%=getadminCartLngStr("LoptDocMCBalLocal")%></option>
				<option <% If myApp.DocMCBal = "D" Then %>selected<% End If %> value="D">
				<%=getadminCartLngStr("DtxtDoc")%></option>
				</select></td>
			</tr>
			<% Else %>
			<input type="hidden" name="DocMCBal" value="<%=myApp.DocMCBal%>">
			<% End If %>
			<tr>
				<td width="435">
				<img src="images/ganchito.gif"><font color="#4783C5"> </font>
				<input type="checkbox" name="SDKLineMemo" value="Y" id="SDKLineMemo" <% If myApp.SDKLineMemo Then %>checked<%end if%> class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="SDKLineMemo"><%=getadminCartLngStr("LtxtSDKLineMemo")%></label></font></td>
				<td>
				<p>
				&nbsp;<select size="1" name="Notes" class="input" id="notes">
				<% 
				GetQuery rd, 6, null, null %>
				<option value="0"><%=getadminCartLngStr("LtxtSelNote")%></option>
				<% do while not rd.eof %>
				<option value="<%=rd("NoteIndex")%>"><%=myHTMLEncode(rd("NoteName"))%></option>
				<% rd.movenext
				loop %>
				</select> <a href="javascript:noteEdit('e')">
				<img height="16" src="images/wpedit.jpg" width="17" border="0" alt="<%=getadminCartLngStr("LtxtEditNote")%>"></a>
				<a href="javascript:noteEdit('a')">
				<img height="13" src="images/newdoc.gif" width="11" align="top" border="0" alt="<%=getadminCartLngStr("LtxtNewNote")%>"></a><a href="javascript:delNote();"><img height="16" src="images/remove.gif" width="16" align="top" border="0" alt="<%=getadminCartLngStr("LtxtDelNote")%>"></a></td>
			</tr>
			<tr>
				<td width="435">
				<img src="images/ganchito.gif"><font face="Verdana" size="1" color="#4783C5"><%=getadminCartLngStr("LtxtEditItemDesQry")%><br><b>select case when exists (...<br>) Then 'Y' Else 'N' End from DOC1, OITM</b></font></td>
				<td>
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td rowspan="2">
							<textarea rows="10" name="ItemDescModQry" dir="ltr" cols="87" class="input" onkeydown="javascript:document.frmAdminCart.btnVerfyFilter.src='images/btnValidate.gif';document.frmAdminCart.btnVerfyFilter.style.cursor = 'pointer';document.frmAdminCart.valItemDescModQry.value='Y';"><% If Not IsNull(myApp.ItemDescModQry) Then %><%=Server.HTMLEncode(myApp.ItemDescModQry)%><% End If %></textarea>
						</td>
						<td valign="top">
							<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminCartLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(25, 'ItemDescModQry', -1, null);">
						</td>
					</tr>
					<tr>
						<td valign="bottom">
							<img src="images/btnValidateDis.gif" id="btnVerfyFilter" alt="<%=getadminCartLngStr("DtxtValidate")%>" onclick="javascript:if (document.frmAdminCart.valItemDescModQry.value == 'Y')VerfyFilter();">	
							<input type="hidden" name="valItemDescModQry" value="N">
						</td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td valign="top"><font size="1" color="#4783C5" face="Verdana"><%=getadminCartLngStr("LtxtAvlVars")%>:</font></td>
				<td>
				<font size="1" color="#4783C5" face="Verdana">
				<span dir="ltr">@SlpCode</span> = <%=getadminCartLngStr("DtxtAgentCode")%><br>
				<span dir="ltr">@ItemCode</span> = <%=getadminCartLngStr("DtxtItemCode")%></font></td>
			</tr>
			</table>
			</div>
		</td>
	</tr>

	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminCartLngStr("DtxtSave")%>" name="B1" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>
<input type="hidden" name="submitCmd" value="adminCart">
</form>
<script type="text/javascript">
<!--
function VerfyFilter()
{
	if (document.frmAdminCart.ItemDescModQry.value != '')
	{
		$.post('verfyQueryFetch.asp', { Type: 'ItemDescQry', Query: document.frmAdminCart.ItemDescModQry.value }, function(data)
		{
			if (data == 'ok')
			{
				Verified();
			}
			else
			{
				alert(data);
			}
		});
	}
	else
	{
		Verified();
	}
}
function Verified()
{
	document.frmAdminCart.btnVerfyFilter.src='images/btnValidateDis.gif'
	document.frmAdminCart.btnVerfyFilter.style.cursor = '';
	document.frmAdminCart.valItemDescModQry.value='N';
}
function changeCartType(value)
{
	<% If Active48 Then %>
	var cmb = document.frmAdminCart.D_DocC;
	var cmbLen = cmb.length;
	switch (value)
	{
		case 'A':
			for (var i = 0;i<cmbLen;i++)
			{
				if (cmb.options[i].value == 'F')
				{
					cmb.remove(i);
					break;
				}
			}
			break;
		case 'S':
			cmb.options[cmbLen] = new Option('<%=getadminCartLngStr("DtxtInvoice")%>/<%=getadminCartLngStr("DtxtReceipt")%>', 'F');
			break;
	}
	<% End If %>
}
//-->
</script>
<!--#include file="bottom.asp" -->