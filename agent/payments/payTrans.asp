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
<!--#include file="lang/payTrans.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="accountControl.asp"-->
<%

           set rs = Server.CreateObject("ADODB.RecordSet")
           set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetPaymentTransferData" & Session("ID")
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
           %>
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getpayTransLngStr("LttlTransPymnt")%></title>
<script language="javascript" src="../generalData.js.asp?dbID=<%=Session("ID")%>&LastUpdate=<%=myApp.LastUpdate%>"></script>
<script language="javascript" src="../general.js"></script>
<script language="javascript">
var DocCur = '<%=rs("DocCur")%>';
var txtValSelAcct = '<%=getpayTransLngStr("DtxtValSelAcct")%>';
var txtValNumVal = "<%=getpayTransLngStr("DtxtValNumVal")%>";
var txtValNumMinVal = "<%=getpayTransLngStr("DtxtValNumMinVal")%>";
var vPagado = '<%=vPagado%>';
var TrsfrSum = '<%=RS("TrsfrSum")%>';
</script>
<script language="javascript" src="payTrans.js"></script>
<script language="javascript" src="accountControl.js"></script>
<!--#include file="../getNumeric.asp"-->
<script type="text/javascript" src="../scr/calendar.js"></script>
<script type="text/javascript" src="../scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="../scr/calendar-setup.js"></script>
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css">
<link rel="stylesheet" type="text/css" media="all" href="../design/0/style/style_cal.css" title="winter">
<script src="http://code.jquery.com/jquery-latest.js"></script>
</head>
<body marginwidth="0" marginheight="0" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" onload="SetSaldo();">
<!--#include file="../licid.inc"-->
<form method="POST" action="submit.asp" name="Form1" onsubmit="return valFrm();">
<div align="center">
	<table border="0" cellpadding="0" width="550" id="table1">
		<tr class="GeneralTlt">
			<td colspan="5"><%=getpayTransLngStr("LttlTransPymnt")%></td>
		</tr>
		<tr class="GeneralTbl">
			<td class="GeneralTblBold2" width="118"><%=getpayTransLngStr("DtxtAccount")%></td>
			<td width="153" colspan="4"><% 
				Dim myAccount
				set myAccount = New AccountControl
				myAccount.ID = "AcctCode"
				myAccount.Value = rs("TrsfrAcct")
				myAccount.DisplayValue = rs("AcctDisp")
				myAccount.Description = rs("AcctName")
				myAccount.AccountType = "cash"
				myAccount.GenerateAccount %></td>
		</tr>
		<tr class="GeneralTbl">
			<td class="GeneralTblBold2" width="118"><%=getpayTransLngStr("LtxtTransDate")%></td>
			<td width="153">
			<input type="text" readonly name="ftrans" id="ftrans" size="29" value="<% If rs("TrsfrDate") <> "" Then response.write FormatDate(rs("TrsfrDate"), False) Else Response.write FormatDate(Now(), False)%>"></td>
			<td width="7">&nbsp;</td>
			<td width="78">&nbsp;</td>
			<td width="182">&nbsp;</td>
		</tr>
		<tr class="GeneralTbl">
			<td class="GeneralTblBold2" width="118"><%=getpayTransLngStr("LtxtRef")%></td>
			<td width="153">
			<input type="text" name="comp" size="29" value="<%=myHTMLEncode(rs("TrsfrRef"))%>" maxlength="11"></td>
			<td width="7">&nbsp;</td>
			<td width="78">&nbsp;</td>
			<td width="182">&nbsp;</td>
		</tr>
		<tr class="GeneralTbl">
			<td class="GeneralTblBold2" width="118">&nbsp;</td>
			<td width="153">&nbsp;</td>
			<td width="7">&nbsp;</td>
			<td class="GeneralTblBold2" width="78"><%=getpayTransLngStr("DtxtTotal")%></td>
			<td width="182">
			<input type="text" name="Total" size="29" value="<%=DocCur%>&nbsp;<%=FormatNumber(rs("TrsfrSum"),myApp.SumDec)%>" onchange="updatePayment(this)" onclick="this.select()" style="text-align: right"></td>
		</tr>
		<tr class="GeneralTbl">
			<td class="GeneralTblBold2" width="118"><%=getpayTransLngStr("DtxtImport")%></td>
			<td width="153">
			<input type="text" readonly class="InputDes" name="ImpInc" size="29" value="<%=DocCur%>&nbsp;<%=FormatNumber(SumApplied,myApp.SumDec)%>" style="text-align: right"></td>
			<td width="7">&nbsp;</td>
			<td width="78">&nbsp;</td>
			<td width="182">&nbsp;</td>
		</tr>
		<tr class="GeneralTbl">
			<td class="GeneralTblBold2" width="118"><%=getpayTransLngStr("LtxtBalToPay")%></td>
			<td width="153">
			<input readonly type="text" class="InputDes" name="SaldoPag" size="29" style="text-align: right"></td>
			<td width="7">&nbsp;</td>
			<td class="GeneralTblBold2" width="78"><%=getpayTransLngStr("DtxtPaid")%></td>
			<td width="182">
			<input type="text" readonly class="InputDes" name="Pagado" size="29" value="<%=DocCur%>&nbsp;<%=FormatNumber(vPagado,myApp.SumDec)%>" style="text-align: right"></td>
		</tr>
		<tr class="GeneralTbl">
			<td class="GeneralTblBold2" width="538" colspan="5">
			<table cellpadding="0" cellspacing="0" border="0" width="100%">
				<tr>
					<td>
					<input type="submit" value="<%=getpayTransLngStr("DtxtAccept")%>" name="Aceptar"></td>
					<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
					<input type="button" name="btnCancel" value="<%=getpayTransLngStr("DtxtCancel")%>" onclick="javascript:if(confirm('<%=getpayTransLngStr("DtxtConfCancel")%>'))window.close();"></td>
				</tr>
			</table></td>
		</tr>
		</table>
</div>
<input type="hidden" name="submitCmd" value="payTrans">
<input type="hidden" name="pagVal" value="0">
<input type="hidden" name="saldofuera" value="<%=myHTMLEncode(Request("saldofuera"))%>">
<input type="hidden" name="AddPath" value="../">
<input type="hidden" name="pop" value="Y">
<input type="hidden" name="DocCur" value="<%=DocCur%>">
</form>
<script type="text/javascript">
    Calendar.setup({
        inputField     :    "ftrans",     // id of the input field
        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
        button         :    "ftrans",  // trigger for the calendar (button ID)
        align          :    "Tr",           // alignment (defaults to "Bl")
        singleClick    :    true
    });
</script>
<div id="clearSpace"></div>
</body>
<% conn.close %>
</html>