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
<!--#include file="lang/payCash.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="accountControl.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<%

           set rs = Server.CreateObject("ADODB.RecordSet")
           If Not Session("PayCart") Then obj = 24 Else obj = 48
           
           set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetPaymentCashData" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@LogNum") = Session("PayRetVal")
			cmd("@OLKObj") = obj
			cmd("@MainCur") = myApp.MainCur
			cmd("@DirectRate") = GetYN(myApp.DirectRate)
			cmd("@SumDec") = myApp.SumDec
           set rs = cmd.execute()
           vPagado = CDbl(rs("Pagado"))
           DocCur = rs("DocCur")
           If rs("YesSaldoFuera") = "Y" Then vPagado = vPagado + (CDbl(rs("SaldoFuera"))*-1)
           SumApplied = CDbl(rs("DocTotal"))
           %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getpayCashLngStr("LtxtCashPymnt")%></title>
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css">
<script language="javascript" src="../generalData.js.asp?dbID=<%=Session("ID")%>&LastUpdate=<%=myApp.LastUpdate%>"></script>
<script language="javascript" src="../general.js"></script>
<script language="javascript" src="accountControl.js"></script>
<script language="javascript">
var DocCur = '<%=DocCur%>';
var vPagado = '<%=vPagado%>';
var cashSum = '<%=RS("CashSum")%>';
var txtValNumVal = "<%=getpayCashLngStr("DtxtValNumVal")%>";
var txtValNumMinVal = "<%=getpayCashLngStr("DtxtValNumMinVal")%>";
</script>
<script language="javascript" src="payCash.js"></script>
<script src="http://code.jquery.com/jquery-latest.js"></script>
<!--#include file="../getNumeric.asp"-->
</head>

<body topmargin="0" leftmargin="0"  onload="SetSaldo()" onfocus="javascript:chkWin();">
<!--#include file="../licid.inc"-->
<form method="POST" name="Form1" action="submit.asp">
<input type="hidden" name="DocCur" value="<%=DocCur%>">
<div align="left">
	<table border="0" cellpadding="0" width="476">
		<tr class="GeneralTlt">
			<td colspan="6"><%=getpayCashLngStr("LtxtCashPymnt")%></td>
		</tr>
		<tr class="GeneralTbl">
			<td width="54"><%=getpayCashLngStr("DtxtAccount")%></td>
			<td width="126" colspan="5">
			<% 
				Dim myAccount
				set myAccount = New AccountControl
				myAccount.ID = "Cuenta"
				myAccount.Value = rs("CashAcct")
				myAccount.DisplayValue = rs("AcctDisp")
				myAccount.Description = rs("AcctName")
				myAccount.AccountType = "cash"
				myAccount.GenerateAccount %>
				</td>
		</tr>
		<tr class="GeneralTbl">
			<td colspan="3">&nbsp;</td>
			<td width="72"><%=getpayCashLngStr("DtxtTotal")%></td>
			<td width="15">
			<img border="0" src="../design/0/images/felcahSelect.gif" width="15" height="13" onclick="SetTotal()"></td>
			<td width="197">
			<input type="text" name="Total" size="31" value="<%=DocCur%>&nbsp;<%=FormatNumber(rs("CashSum"),myApp.SumDec)%>" onchange="chkNum(this);updatePayment()" onfocus="this.select()" style="text-align: right"></td>
		</tr>
		<tr class="GeneralTbl">
			<td width="54"><%=getpayCashLngStr("DtxtImport")%></td>
			<td width="126">
			<input class="InputDes" readonly type="text" name="ImpInc" size="24" value="<%=DocCur%>&nbsp;<%=FormatNumber(SumApplied,myApp.SumDec)%>" style="text-align: right"></td>
			<td width="276" colspan="4">&nbsp;</td>
		</tr>
		<tr class="GeneralTbl">
			<td width="54"><%=getpayCashLngStr("LtxtBalToPay")%></td>
			<td width="126">
			<input class="InputDes" readonly type="text" name="SaldoPag" size="24" style="text-align: right"></td>
			<td width="72" colspan="2"><%=getpayCashLngStr("DtxtPaid")%></td>
			<td width="214" colspan="2">
			<input type="text" class="InputDes" readonly name="Pagado" size="35" value="<%=DocCur%>&nbsp;<%=FormatNumber(vPagado,myApp.SumDec)%>" style="text-align: right"></td>
		</tr>
		<tr class="GeneralTbl">
			<td colspan="6">
			<table cellpadding="0" cellspacing="0" border="0" width="100%">
				<tr>
					<td>
					<input type="submit" value="<%=getpayCashLngStr("DtxtAccept")%>" name="btnAccept"></td>
					<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
					<input type="button" name="btnCancel" value="<%=getpayCashLngStr("DtxtCancel")%>" onclick="javascript:if(confirm('<%=getpayCashLngStr("DtxtConfCancel")%>'))window.close();"></td>
				</tr>
			</table></td>
		</tr>
	</table>
</div>
<input type="hidden" name="pagVal" value="0">
<input type="hidden" name="submitCmd" value="payCash">
<input type="hidden" name="saldofuera" value="<%=Request("saldofuera")%>">
<input type="hidden" name="AddPath" value="../">
<input type="hidden" name="pop" value="Y">
</form>
<div id="clearSpace"></div>
</body>
<% conn.close %>
</html>