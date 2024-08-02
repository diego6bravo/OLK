<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="../lcidReturn.inc"-->
<!--#include file="lang/payCred.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<%

           set rs = Server.CreateObject("ADODB.RecordSet")
           set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetPaymentCreditData" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@LogNum") = Session("PayRetVal")
			cmd("@MainCur") = myApp.MainCur
			cmd("@DirectRate") = GetYN(myApp.DirectRate)
			cmd("@SumDec") = myApp.SumDec
			set rs = cmd.execute()
           vPagado = CDbl(rs("Pagado"))
           DocCur = rs("DocCur")
           If rs("YesSaldoFuera") = "Y" Then vPagado = vPagado + (CDbl(rs("SaldoFuera"))*-1)
           SumApplied = CDbl(rs("DocTotal"))
           
           set rt = Server.CreateObject("ADODB.Recordset")
           set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetPaymentCreditDetails" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@LogNum") = Session("PayRetVal")
           rt.open cmd, , 3, 1
           
           %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getpayCredLngStr("LttlCredCardPymnt")%></title>
<script language="javascript" src="../generalData.js.asp?dbID=<%=Session("ID")%>&LastUpdate=<%=myApp.LastUpdate%>"></script>
<script language="javascript" src="../general.js"></script>
<!--#include file="../getNumeric.asp"-->
<script language="javascript">
var DocCur = "<%=DocCur%>";
var txtValNumVal = "<%=getpayCredLngStr("DtxtValNumVal")%>";
var txtValNumMinVal = "<%=getpayCredLngStr("DtxtValNumMinVal")%>";
var txtValNumMaxVal = "<%=getpayCredLngStr("DtxtValNumMaxVal")%>";
var voucher = '<%=Request("voucher")%>';
var Pagado = '<%=rs("Pagado")%>';
var creditsum = '<%=RS("creditsum")%>';
var txtValPymntSys = "<%=getpayCredLngStr("LtxtValPymntSys")%>";
var txtValCardNoPymnt = "<%=getpayCredLngStr("LtxtValCardNoPymnt")%>";
var txtValCardDueDat = "<%=getpayCredLngStr("LtxtValCardDueDat")%>";
var txtValAut = "<%=getpayCredLngStr("LtxtValAut")%>";
var txtMinCredPymnt = "<%=getpayCredLngStr("LtxtMinCredPymnt")%>";
</script>
<% If Request("voucher") = "add" Then %>
<script language="javascript" src="payCredAdd.js"></script>
<% ElseIf IsNumeric(Request("voucher")) Then %>
<script language="javascript" src="payCredUpd.js"></script>
<% End If %>
<script language="javascript" src="payCred.js"></script>
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css">
</head>
<body marginwidth="0" marginheight="0" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" onload="javascript:SetSaldo();<% If Request("voucher") = "add" Then%> Start('cards.asp','300','200','auto')<% end if %>" onfocus="javascript:chkWin();">
<!--#include file="../licid.inc"-->
<div align="center">
	<table border="0" cellpadding="0" width="100%">
		<tr class="GeneralTlt">
			<td><%=getpayCredLngStr("LttlCredCardPymnt")%></td>
		</tr>
		<form method="POST" action="payCred.asp" name="FormTop">
		<input type="hidden" name="pop" value="Y">
		<input type="hidden" name="AddPath" value="../">
		<tr class="GeneralTbl">
			<td>
			<p align="right"><%=getpayCredLngStr("LtxtCard")%> <select size="1" name="voucher" onchange="submit()">
			<option value="NULL"><%=getpayCredLngStr("LtxtDoSel")%></option>
			<option <% If Request("voucher") = "add" then response.write "selected "%>value="add">
			<%=getpayCredLngStr("LtxtAddCard")%></option>
			<% if not rt.eof then
			while not rt.eof %>
			<option <% If CStr(Request("voucher")) = CStr(rt("LineNum")) then response.write "selected "%>value="<%=rt("LineNum")%>"><%=rt("LineNum")+1%>-<%=myHTMLEncode(rt("CardName"))%></option>
			<% rt.movenext
			wend 
			rt.movefirst
			end if %>
			</select></td>
		</tr>
		<input type="hidden" name="imp" value="<%=myHTMLEncode(Request("imp"))%>">
		<input type="hidden" name="saldofuera" value="<%=myHTMLEncode(Request("saldofuera"))%>">
		</form>
<% If Request("voucher") = "add" Then %>
		<form method="POST" action="submit.asp" name="Form1" onsubmit="return valFrm();">
		<tr class="GeneralTbl">
			<td>
			<table border="0" cellpadding="0" width="100%">
				<tr class="GeneralTblBold2">
					<td align="center" width="123"><%=getpayCredLngStr("LtxtCard")%></td>
					<td align="center"><%=getpayCredLngStr("DtxtCode")%></td>
					<td align="center"><%=getpayCredLngStr("LtxtCardNum")%></td>
					<td align="center"><%=getpayCredLngStr("LtxtCardValUn")%></td>
					<td align="center"><%=getpayCredLngStr("LtxtID")%></td>
				</tr>
				<tr>
					<td width="123" height="18">
					<p align="center">
					<input readonly type="text" name="tarjeta" size="17" onclick="Start('cards.asp','300','200','auto')"></td>
					<td height="18"><input type="hidden" name="CreditAcct"><input readonly  class="InputDes" type="text" name="CreditAcctDisp" size="17" value=""></td>
					<td height="18">
					<p align="left">
					<input type="text" name="CrCardNum" size="17" maxlength="20"></td>
					<td height="18"><input type="text" name="CardValidM" size="4" onchange="if(this.value!='MM'){chkNum(this, 12, 1, '')}else{this.value=''};" onfocus="this.select();" maxlength="2" value="MM" onclick="Ifthis(this,'MM')">
					<input type="text" name="CardValidY" size="4" onchange="if(this.value!='YY'){chkNum(this,99,0, '')}else{this.value=''};" onfocus="this.select();" maxlength="2" value="YY" onclick="Ifthis(this,'YY')"></td>
					<td height="18">
					<p align="center"><input type="text" name="OwnerIdNum" size="17" maxlength="15"></td>
				</tr>
				<tr class="GeneralTblBold2">
					<td width="123" align="center"><%=getpayCredLngStr("DtxtPhone")%></td>
					<td align="center"><%=getpayCredLngStr("LtxtPymntSystem")%></td>
					<td align="center"><%=getpayCredLngStr("DtxtImport")%></td>
					<td align="center"><%=getpayCredLngStr("LtxtNumOfPay")%></td>
					<td align="center"><%=getpayCredLngStr("LtxtFirstPymnt")%></td>
				</tr>
				<tr>
					<td width="123" align="center">
					<input type="text" name="OwnerPhone" size="17" maxlength="20"></td>
					<td align="center">
					<input readonly type="text" name="syspag" size="17" onclick="StartSisPag(), this.blur()"></td>
					<td align="center">
					<input name="impval" size="17" onchange="javascript:chkNum(this,null,null, ''); chkVal(this); optChangeImp('N')" onclick="this.select()" style="float: left; text-align:right"></td>
					<td align="center"><input class="InputDes" readonly type="text" name="pagcant" size="17" value="1" onchange="chkNum(this,null,1,1);chkVal(this); optChangeImp('Y')" onclick="this.select()"></td>
					<td align="center">
					<input class="InputDes" readonly type="text" name="perpago" size="17" onclick="this.select()" onchange="chkNum(this,null,null,var1); chkVal(this); optChangePerPago()" style="text-align: right"></td>
				</tr>
				<tr class="GeneralTblBold2">
					<td width="123" align="center"><%=getpayCredLngStr("LtxtAddPayQty")%>l</td>
					<td align="center"><%=getpayCredLngStr("LtxtRefNum")%></td>
					<td align="center"><%=getpayCredLngStr("LtxtAut")%></td>
					<td align="center">&nbsp;</td>
					<td align="center">&nbsp;</td>
				</tr>
				<tr>
					<td width="123">
					<p align="center">
					<input readonly class="InputDes" type="text" name="cpagoadd" size="17" onclick="this.select()" onchange="chkNum(this,null,null,var2); chkVal(this); optChangeCPago()" style="text-align: right"></td>
					<td>
					<input type="text" name="compnum" size="17" maxlength="20"></td>
					<td colspan="2">
					<input type="text" name="autorizacion" size="46" maxlength="20"></td>
					<td>
					<p align="center">
					&nbsp;</td>
				</tr>
				<tr>
					<td width="123">
					<input type="submit" value="<%=getpayCredLngStr("DtxtAdd")%>" name="addcard1"></td>
					<td>
					&nbsp;</td>
					<td colspan="2">
					&nbsp;</td>
					<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
					<input type="button" name="btnCancel" value="<%=getpayCredLngStr("DtxtCancel")%>" onclick="javascript:if(confirm('<%=getpayCredLngStr("DtxtConfCancel")%>')){document.FormTop.voucher.selectedIndex=0;document.FormTop.submit();}"></td>
				</tr>
			</table>
			</td>
		</tr>
		<input type="hidden" name="submitCmd" value="addCred">
			<input type="hidden" name="CreditCard" value="0">
			<input type="hidden" name="SistPagCode" value="0">
			<input type="hidden" name="imp" value="<%=myHTMLEncode(Request("imp"))%>">
			<input type="hidden" name="saldofuera" value="<%=myHTMLEncode(Request("saldofuera"))%>">
			<input type="hidden" name="AddPath" value="../">
			<input type="hidden" name="pop" value="Y">
			<input type="hidden" name="DocCur" value="<%=DocCur%>">
		</form>
		<% ElseIf IsNumeric(Request("voucher")) Then 
		set rc = server.createobject("ADODB.Recordset")
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetPaymentCreditDetailsData" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@LogNum") = Session("PayRetVal")
		cmd("@LineNum") = Request("VOUCHER")
		set rc = cmd.execute() %>
		<script language="javascript">
		InstalMent = '<%=rc("InstalMent")%>';
		</script>
		<form method="POST" action="submit.asp" name="Form3" onsubmit="return valFrm();">
		<tr class="GeneralTbl">
			<td>
			<table border="0" cellpadding="0" width="100%">
				<tr class="GeneralTblBold2">
					<td align="center" width="123"><%=getpayCredLngStr("LtxtCard")%></td>
					<td align="center"><%=getpayCredLngStr("DtxtCode")%></td>
					<td align="center"><%=getpayCredLngStr("LtxtCardNum")%></td>
					<td align="center"><%=getpayCredLngStr("LtxtCardValUn")%></td>
					<td align="center"><%=getpayCredLngStr("LtxtID")%></td>
				</tr>
				<tr>
					<td width="123">
					<p align="center">
					<input readonly type="text" name="tarjeta" size="17" onclick="Start('cards.asp','300','200','auto')" value="<%=myHTMLEncode(rc("CardName"))%>"></td>
					<td>
					<input type="hidden" name="CreditAcct" value="<%=myHTMLEncode(rc("CreditAcct"))%>"><input readonly  class="InputDes" type="text" name="CreditAcctDisp" size="17" value="<%=myHTMLEncode(rc("AcctDisp"))%>"></td>
					<td>
					<input type="text" name="CrCardNum" size="17" maxlength="20" value="<%=myHTMLEncode(rc("CrCardNum"))%>"></td>
					<td><input type="text" name="CardValidM" size="4" onchange="chkNum(this,12,1, <%=rc("CardValidM")%>);" onfocus="this.select();" maxlength="2" value="<%=rc("CardValidM")%>">
					<input type="text" name="CardValidY" size="4" onchange="chkNum(this,99,0, <%=rc("CardValidY")%>);" onfocus="this.select();" maxlength="2" value="<%=rc("CardValidY")%>"></td>
					<td>
					<p align="center"><input type="text" name="OwnerIdNum" size="17" maxlength="15" value="<%=myHTMLEncode(rc("OwnerIdNum"))%>"></td>
				</tr>
				<tr class="GeneralTblBold2">
					<td width="123" align="center"><%=getpayCredLngStr("DtxtPhone")%></td>
					<td align="center"><%=getpayCredLngStr("LtxtPymntSystem")%></td>
					<td align="center"><%=getpayCredLngStr("DtxtImport")%></td>
					<td align="center"><%=getpayCredLngStr("LtxtNumOfPay")%></td>
					<td align="center"><%=getpayCredLngStr("LtxtFirstPymnt")%></td>
				</tr>
				<tr>
					<td width="123" align="center">
					<input type="text" name="OwnerPhone" size="17" maxlength="20" value="<%=rc("OwnerPhone")%>"></td>
					<td align="center">
					<input readonly type="text" name="syspag" size="17" onclick="StartSisPag()" value="<%=myHTMLEncode(rc("CrTypeName"))%>"></td>
					<td align="center">
					<p align="left">
					<input type="text" name="impval" size="17" onchange="chkNum(this,null,null, '<%=DocCur%> <%=FormatNumber(rc("CreditSum"),myApp.SumDec)%>'); chkVal(this); optChangeImp()" onclick="this.select()" value="<%=DocCur%>&nbsp;<%=FormatNumber(rc("CreditSum"),myApp.SumDec)%>" style="text-align: right"></td>
					<td align="center"><input type="text" <% If rc("InstalMent") = "N" Then %>readonly  class="InputDes"<% end if %> name="pagcant" size="17" value="<%=rc("NumOfPmnts")%>" onchange="chkVal(this); optChangeImp()" onclick="this.select()"></td>
					<td align="center">
					<input type="text" <% If rc("InstalMent") = "N" Then %>readonly  class="InputDes"<%end if%> name="perpago" size="17" onclick="this.select()" onchange="chkNum(this,null,null,var1); chkVal(this); optChangePerPago()" value="<%=DocCur%>&nbsp;<%=FormatNumber(rc("FirstSum"),myApp.SumDec)%>" style="text-align: right"></td>
				</tr>
				<tr class="GeneralTblBold2">
					<td width="123" align="center"><%=getpayCredLngStr("LtxtAddPayQty")%></td>
					<td align="center"><%=getpayCredLngStr("LtxtRefNum")%></td>
					<td align="center"><%=getpayCredLngStr("LtxtAut")%></td>
					<td align="center">&nbsp;</td>
					<td align="center">&nbsp;</td>
				</tr>
				<tr>
					<td width="123">
					<p align="center">
					<input <% If rc("InstalMent") = "N" Then %>readonly  class="InputDes"<%end if%> type="text" name="cpagoadd" size="17" onclick="this.select()" onchange="chkNum(this,null,null,var2); chkVal(this); optChangeCPago()" value="<% If Not ISNUll(rc("AddPmntSum")) Then %><%=DocCur%>&nbsp;<%=FormatNumber(rc("AddPmntSum"),myApp.SumDec)%><% End If %>" style="text-align: right"></td>
					<td>
					<input type="text" name="compnum" size="17" maxlength="20" value="<%=myHTMLEncode(rc("VoucherNum"))%>"></td>
					<td colspan="2">
					<input type="text" name="autorizacion" size="46" maxlength="20" value="<%=myHTMLEncode(rc("ConfNum"))%>"></td>
					<td>
					<p align="center">
					&nbsp;</td>
				</tr>
				<tr>
					<td width="123">
					<input type="submit" value="<%=getpayCredLngStr("DtxtUpdate")%>" name="addcard"></td>
					<td>
					&nbsp;</td>
					<td colspan="2">
					&nbsp;</td>
					<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
					<input type="button" name="btnCancel" value="<%=getpayCredLngStr("DtxtCancel")%>" onclick="javascript:if(confirm('<%=getpayCredLngStr("DtxtConfCancel")%>')){document.FormTop.voucher.selectedIndex=0;document.FormTop.submit();}"></td>
				</tr>
			</table>
			</td>
		</tr>
		<input type="hidden" name="submitCmd" value="updateCred">
		<input type="hidden" name="CreditCard" value="<%=rc("CreditCard")%>">
		<input type="hidden" name="SistPagCode" value="<%=rc("CrTypeCode")%>">
		<input type="hidden" name="imp" value="<%=myHTMLEncode(Request("imp"))%>">
		<input type="hidden" name="linenum" value="<%=Request("voucher")%>">
		<input type="hidden" name="saldofuera" value="<%=myHTMLEncode(Request("saldofuera"))%>">
		<input type="hidden" name="AddPath" value="../">
		<input type="hidden" name="pop" value="Y">
		<input type="hidden" name="DocCur" value="<%=DocCur%>"></form>
		<% end if 
		If not rt.eof then %>
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%">
				<tr class="GeneralTblBold2">
					<td align="center">#</td>
					<td align="center"><%=getpayCredLngStr("LtxtCard")%></td>
					<td align="center"><%=getpayCredLngStr("LtxtDueDate")%></td>
					<td align="center"><%=getpayCredLngStr("DtxtTotal")%></td>
					<td align="center" width="27">&nbsp;</td>
				</tr>
				<% do while not rt.eof %>
				<tr class="GeneralTbl">
					<td align="center"><%=rt("LineNum")+1%>&nbsp;</td>
					<td align="center"><%=rt("CardName")%>&nbsp;</td>
					<td align="center"><%=FormatDate(rt("CardValid"), True)%>&nbsp;</td>
					<td align="right"><%=DocCur%>&nbsp;<%=FormatNumber(rt("CreditSum"),myApp.SumDec)%></td>
					<td width="27" align="center">
					<p align="center">
					<a href="javascript:if(confirm('<%=getpayCredLngStr("LtxtConfDelCard")%>'.replace('{0}', '<%=rt("CardName")%>')))doMyLink('submit.asp', 'submitCmd=delCard&linenum=<%=rt("linenum")%>&imp=<%=Request("imp")%>&saldofuera=<%=Request("saldofuera")%>', '');else">
					<img border="0" src="../design/0/images/<%=Session("rtl")%>xicon.gif" width="12" height="11"></a></td>
				</tr>
				<% rt.movenext
				loop %>
			</table>
			</td>
		</tr>
		<% end if %>
		<form method="POST" action="submit.asp" name="Form2">
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%" cellspacing="1">
				<tr class="GeneralTbl">
					<td align="center" width="79">&nbsp;</td>
					<td align="center">&nbsp;</td>
					<td align="center" width="39"><%=getpayCredLngStr("DtxtTotal")%></td>
					<td align="center" width="106">
					<input readonly class="InputDes" type="text" name="Total" size="34" value="<%=DocCur%>&nbsp;<%=FormatNumber(rs("creditsum"),myApp.SumDec)%>" onchange="updatePayment()" onclick="this.select()" style="text-align: right"></td>
				</tr>
				<tr class="GeneralTbl">
					<td width="79"><%=getpayCredLngStr("DtxtImport")%></td>
					<td align="center">
					<input readonly class="InputDes" name="ImpInc" size="39" style="float: left; text-align:right" value="<%=DocCur%>&nbsp;<%=FormatNumber(sumapplied,myApp.SumDec)%>"></td>
					<td align="center" width="39">&nbsp;</td>
					<td align="center" width="106">&nbsp;</td>
				</tr>
				<tr class="GeneralTbl">
					<td width="79"><%=getpayCredLngStr("LtxtBalToPay")%></td>
					<td align="center">
					<input readonly class="InputDes" name="SaldoPag" size="39" style="float: left; text-align:right"></td>
					<td align="center" width="39"><%=getpayCredLngStr("DtxtPaid")%></td>
					<td align="center" width="106">
					<input readonly class="InputDes" type="text" name="Pagado" size="34" style="text-align: right" value="<%=DocCur%>&nbsp;<%=FormatNumber(vPagado,myApp.SumDec)%>"></td>
				</tr>
				<tr class="GeneralTbl">
					<td align="center" colspan="4">
					<table cellpadding="0" cellspacing="0" border="0" width="100%">
						<tr>
							<td align="center">
							&nbsp;</td>
						</tr>
						<tr>
							<td align="center">
							<input type="submit" value="<%=getpayCredLngStr("DtxtClose")%>" name="aceptar"></td>
						</tr>
						<tr>
							<td align="center">
							&nbsp;</td>
						</tr>
					</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
			<input type="hidden" name="submitCmd" value="payCred">
			<input type="hidden" name="pagVal" value="0">
			<input type="hidden" name="saldofuera" value="<%=Request("saldofuera")%>">
		</form>
		</table>
</div>
<!--#include file="../linkForm.asp"-->
</body>
<% conn.close %></html>