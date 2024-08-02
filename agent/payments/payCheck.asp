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
<!--#include file="lang/payCheck.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="accountControl.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<%
 
set rs = Server.CreateObject("ADODB.RecordSet")
If Not Session("PayCart") Then obj = 24 Else obj = 13
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetPaymentCheckData" & Session("ID")
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
SumApplied = CDbl(rs("DocTotal"))
If rs("YesSaldoFuera") = "Y" Then vPagado = vPagado + (CDbl(rs("SaldoFuera"))*-1)

set rb = Server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetBanksList" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
rb.open cmd, , 3, 1

set rsL = server.createobject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetPaymentCheckDetails"
cmd.Parameters.Refresh()
cmd("@LogNum") = Session("PayRetVal")
rsL.open cmd, , 3, 1
%>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getpayCheckLngStr("LttlChkPymnt")%></title>
<script language="javascript" src="../generalData.js.asp?dbID=<%=Session("ID")%>&LastUpdate=<%=myApp.LastUpdate%>"></script>
<script language="javascript" src="../general.js"></script>
<script language="javascript">
var DocCur = '<%=DocCur%>';
var txtValNumVal = "<%=getpayCheckLngStr("DtxtValNumVal")%>";
var txtValNumMinVal = "<%=getpayCheckLngStr("DtxtValNumMinVal")%>";
var txtValChkImp = "<%=getpayCheckLngStr("LtxtValChkImp")%>";
var txtConfDelChk = '<%=getpayCheckLngStr("LtxtConfDelChk")%>';
var vPagado = '<%=vPagado%>';
var checkSum = '<%=RS("checkSum")%>';
</script>
<script language="javascript" src="payCheck.js"></script>
<script type="text/javascript" src="../scr/calendar.js"></script>
<script type="text/javascript" src="../scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="../scr/calendar-setup.js"></script>
<script language="javascript" src="accountControl.js"></script>
<!--#include file="../getNumeric.asp"-->
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css">
<link rel="stylesheet" type="text/css" media="all" href="../design/0/style/style_cal.css" title="winter">
<script src="http://code.jquery.com/jquery-latest.js"></script>

</head>

<body topmargin="0" leftmargin="0" onload="SetSaldo()" onfocus="javascript:chkWin();">
<!--#include file="../licid.inc"-->
<form method="POST" action="submit.asp" name="Form1">
<input type="hidden" name="DocCur" value="<%=DocCur%>">
<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTlt">
		<td colspan="2"><%=getpayCheckLngStr("LttlChkPymnt")%></td>
	</tr>
	<tr class="GeneralTbl">
		<td class="GeneralTblBold2"><%=getpayCheckLngStr("DtxtAccount")%></td>
		<td><% 
				Dim myAccount
				set myAccount = New AccountControl
				myAccount.ID = "checkAcct"
				myAccount.Value = rs("checkAcct")
				myAccount.DisplayValue = rs("AcctDisp")
				myAccount.Description = rs("AcctName")
				myAccount.AccountType = "check"
				myAccount.GenerateAccount %></td>
	</tr>
	<tr>
		<td colspan="2">
		<table border="0" cellpadding="0" width="100%">
			<tr class="GeneralTblBold2">
				<td align="center"><%=getpayCheckLngStr("DtxtDate")%></td>
				<td align="center"><%=getpayCheckLngStr("LtxtBank")%></td>
				<td align="center"><%=getpayCheckLngStr("DtxtBranch")%></td>
				<td align="center"><%=getpayCheckLngStr("DtxtAccount")%></td>
				<td align="center"><%=getpayCheckLngStr("LtxtCheck")%></td>
				<td align="center"><%=getpayCheckLngStr("DtxtImport")%></td>
				<td align="center">&nbsp;</td>
			</tr>
			<% do while not rsL.eof %>
			<tr class="GeneralTbl">
				<td><input type="hidden" name="LineNum" value="<%=rsL("LineNum")%>">
				<p align="center">
				<input readonly type="text" name="fecha<%=rsL("LineNum")%>" id="fecha<%=rsL("LineNum")%>" size="9" value="<%=FormatDate(rsL("DueDate"), False)%>"></td>
				<td>
				<select size="1" name="banco<%=rsL("LineNum")%>" style="width: 200px;">
				<option></option>
				<% do while not rb.eof %>
				<option <% If Not IsNull(rsL("bankcode")) Then %><% If CStr(rb("bankcode")) = CStr(rsL("bankcode")) then response.write "selected " %><% End If %>value="<%=rb("bankCode")%>"><%=myHTMLEncode(rb("bankName"))%></option>
				<% rb.movenext
				loop
				rb.movefirst %>
				</select></td>
				<td>
				<p align="center">
				<input type="text" name="sucursal<%=rsL("LineNum")%>" size="9" value="<%=myHTMLEncode(rsL("Branch"))%>" maxlength="50" onchange="javascript:changeBranch('<%=rsL("LineNum")%>', this.value);"></td>
				<td>
				<p align="center">
				<input type="text" name="cuenta<%=rsL("LineNum")%>" size="11" value="<%=myHTMLEncode(rsL("AcctNum"))%>" maxlength="50"></td>
				<td>
				<p align="center">
				<input type="text" name="detalles<%=rsL("LineNum")%>" size="8" value="<%=myHTMLEncode(rsL("CheckNum"))%>" maxlength="100" onchange="chkCheck(this);"></td>
				<td>
				<p align="center">
				<input type="text" name="imp<%=rsL("LineNum")%>" id="imp<%=rsL("LineNum")%>" size="20" value="<%=DocCur%>&nbsp;<%=FormatNumber(rsL("CheckSum"),myApp.SumDec)%>" onchange="chkNum(this, 0);setTotal(this, '<%=rsL("CheckSum")%>')" maxlength="40" onfocus="javascript:this.select()" style="text-align: right"></td>
				<td>
				<p align="center">
				<a href="javascript:delCheck(<%=rsL("LineNum")%>);">
				<img border="0" src="../design/0/images/<%=Session("rtl")%>xicon.gif" width="12" height="11"></a></td>
			</tr>
          <% rsL.movenext
          loop %>
			<tr class="GeneralTbl">
				<td>
				<p align="center">
				<input readonly type="text" name="fecha" id="fecha" size="9" value="<%If Request("fecha") = "" Then%><%=FormatDate(Now(), False)%><%else%><%=Request("fecha")%><%end if%>"></td>
				<td><select size="1" name="banco" style="width: 200">
				<option></option>
				<% 
				If Request("banco") <> "" Then
					selBanco = Request("banco")
				Else
					selBanco = rs("BankCode")
				End If
				do while not rb.eof %>
				<option <% If selBanco = rb("bankcode") then response.write "selected "%>value="<%=rb("bankCode")%>"><%=myHTMLEncode(rb("bankName"))%></option>
				<% rb.movenext
				loop %>
				</select></td>
				<td>
				<p align="center">
				<input type="text" name="sucursal" size="9" value="<%=myHTMLEncode(selBranch)%>" maxlength="50" onchange="javascript:changeBranch('new', this.value);"></td>
				<td>
				<p align="center">
				<input type="text" name="cuenta" size="11" value="<%=myHTMLEncode(selAcct)%>" maxlength="50"></td>
				<td>
				<p align="center">
				<input type="text" name="detalles" size="8" value="<%=myHTMLEncode(Request("detalles"))%>" onchange="chkCheck(this);"></td>
				<td>
				<p align="center">
				<input type="text" name="impval" size="20" value="<%=Request("impval")%>" style="text-align: right"  maxlength="40" onchange="chkNum(this, 0);setTotal(this, '')"></td>
				<td>&nbsp;</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr class="GeneralTbl">
		<td colspan="2"><hr size="1"></td>
	</tr>
	<tr class="GeneralTbl">
		<td colspan="2">
		<table border="0" cellpadding="0" width="100%" id="table3">
			<tr>
				<td colspan="2">&nbsp;</td>
				<td class="GeneralTblBold2"><%=getpayCheckLngStr("DtxtTotal")%></td>
				<td class="GeneralTblBold2">
				<input readonly class="InputDes" type="text" name="Total" size="20" value="<%=DocCur%>&nbsp;<%=FormatNumber(rs("checkSum"),myApp.SumDec)%>" style="text-align: right"></td>
			</tr>
			<tr>
				<td class="GeneralTblBold2"><%=getpayCheckLngStr("DtxtImport")%></td>
				<td class="GeneralTblBold2">
				<input readonly class="InputDes" type="text" name="ImpInc" size="20" value="<%=DocCur%>&nbsp;<%=FormatNumber(SumApplied,myApp.SumDec)%>" style="text-align: right"></td>
				<td colspan="2" class="GeneralTblBold2">&nbsp;</td>
			</tr>
			<tr>
				<td class="GeneralTblBold2"><%=getpayCheckLngStr("LtxtBalToPay")%></td>
				<td class="GeneralTblBold2">
				<input readonly class="InputDes" type="text" name="SaldoPag" size="20" style="text-align: right"></td>
				<td class="GeneralTblBold2"><%=getpayCheckLngStr("DtxtPaid")%></td>
				<td class="GeneralTblBold2">
				<input readonly class="InputDes" type="text" name="Pagado" size="20" value="<%=DocCur%>&nbsp;<%=FormatNumber(vPagado,myApp.SumDec)%>" style="text-align: right"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr class="GeneralTbl">
		<td colspan="2">
			<table cellpadding="0" cellspacing="0" border="0" width="100%">
				<tr>
					<td>
					<input type="button" value="<%=getpayCheckLngStr("LtxtAddCheck")%>" name="B1" onclick="chkNewChk()"> -
					<input type="submit" value="<%=getpayCheckLngStr("DtxtAccept")%>" name="Aceptar"></td>
					<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
					<input type="button" name="btnCancel" value="<%=getpayCheckLngStr("DtxtCancel")%>" onclick="javascript:if(confirm('<%=getpayCheckLngStr("DtxtConfCancel")%>'))window.close();"></td>
				</tr>
			</table>
		</td>
	</tr>
	</table>
<input type="hidden" name="pagVal" value="0">
<input type="hidden" name="imp" value="<%=myHTMLEncode(Request("imp"))%>">
<input type="hidden" name="submitCmd" value="payCheck">
<input type="hidden" name="Agregar" value="">
<input type="hidden" name="saldofuera" value="<%=myHTMLEncode(Request("saldofuera"))%>">
<input type="hidden" name="AddPath" value="../">
<input type="hidden" name="pop" value="Y">
</form>

&nbsp;
<script type="text/javascript">
    Calendar.setup({
        inputField     :    "fecha",     // id of the input field
        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
        button         :    "fecha",  // trigger for the calendar (button ID)
        align          :    "Bl",           // alignment (defaults to "Bl")
        singleClick    :    true
    });
    <% If rsL.recordcount > 0 then rsL.movefirst
    do while not rsL.eof %>
    Calendar.setup({
        inputField     :    "fecha<%=rsL("LineNum")%>",     // id of the input field
        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
        button         :    "fecha<%=rsL("LineNum")%>",  // trigger for the calendar (button ID)
        align          :    "Bl",           // alignment (defaults to "Bl")
        singleClick    :    true
    });
    <% rsL.movenext
    loop %>
</script>
<div id="clearSpace"></div>
<!--#include file="../linkForm.asp"-->
</body>

</html>
<% set rs = nothing
set rsL = nothing
set rb = nothing
conn.close %>