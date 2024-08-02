<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Not (Session("UserName") <> "" and CardType = "C" and myApp.EnableORCT) Then Response.Redirect "unauthorized.asp" %>
<% addLngPathStr = "" %>
<!--#include file="lang/agentPayment.asp" -->
<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>
<!--#include file="lcidReturn.inc"-->

<% 

set rc = Server.CreateObject("ADODB.recordset")

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetPaymentData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@LogNum") = Session("PayRetVal")
cmd("@SlpCode") = Session("vendid")
cmd("@MainCur") = myApp.MainCur
cmd("@DirectRate") = GetYN(myApp.DirectRate)
cmd("@SumDec") = myApp.SumDec
rc.open cmd, , 3, 1
vPagado = CDbl(rc("Pagado"))
DocDate = rc("DocDate")
If rc("YesSaldoFuera") = "Y" Then vPagado = vPagado + (CDbl(rc("SaldoFuera"))*-1)
chkOpt = rc("EnableSDK") = "Y"

set rs = Server.CreateObject("ADODB.recordset")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetPaymentLinesData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LogNum") = Session("PayRetVal")
cmd("@CardCode") = rc("CardCode")
cmd("@MainCur") = myApp.MainCur
cmd("@ApplyOpenRctToInvBal") = GetYN(myApp.ApplyOpenRctToInvBal)
rs.open cmd, , 3, 1

set rd = Server.CreateObject("ADODB.recordset")
GetQuery rd, 4, 24, null

set rx = Server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetBPContacts" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@CardCode") = Session("UserName")
rx.open cmd, , 3, 1

DocCur = rc("DocCur")
ShowCurncy = rc("ShowCurncy")
EnableMC = rc("EnableMC")
%>
<script language="javascript">
var selDes = '<%=SelDes%>';
var dbName = '<%=Session("olkdb")%>';
var txtErrSaveData = '<%=getagentPaymentLngStr("DtxtErrSaveData")%>';
var MainCur = '<%=myApp.MainCur%>';
var DocCur = '<%=DocCur%>';
var txtValFldMaxChar = '<%=getagentPaymentLngStr("DtxtValFldMaxChar")%>';
var txtValNumVal = '<%=getagentPaymentLngStr("DtxtValNumVal")%>';
var txtValNumValWhole = '<%=getagentPaymentLngStr("DtxtValNumValWhole")%>';
</script>
<script type="text/javascript" src="payments/payment.js"></script>
<script language="javascript">
function chkSubmit()
{
<% 
set rcOpt = Server.CreateObject("ADODB.RecordSet")
cmd.CommandText = "DBOLKGetUDFNotNull" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@UserType") = "V"
cmd("@TableID") = "ORCT"
cmd("@OP") = "O"
rcOpt.open cmd, , 3, 1
Do while not rcOpt.eof %>
	if (document.Form1.U_<%=rcOpt("AliasID")%>.value == "")
	{
		alert('<%=getagentPaymentLngStr("LtxtValFld")%>'.replace('{0}', '<%=rcOpt("Descr")%>'));
		document.Form1.U_<%=rcOpt("AliasID")%>.focus();
		return false;
	}
<% 
rcOpt.movenext
loop
if rcOpt.recordcount > 0 then rcOpt.movefirst %>
	if (!noMsg)
	{
		var pagado = parseFloat(getNumericVB(document.Form1.pagado.value.replace(DocCur, '')));
		var importe = parseFloat(getNumericVB(document.Form1.importe.value.replace(DocCur, '')));
		if (pagado == importe)
		{
			if (confirm('<%=getagentPaymentLngStr("LtxtConfAddDoc")%>'))
			{
				document.Form1.finish.value = 'Y';
			}
			else
			{
				return false;
			}
		}
		else
		{
			var varSaldo = OLKFormatNumber(pagado - importe, SumDec);
			<% 
			txtConfVar0 = LtxtDeLa & " " & LCase(txtInv) %>
			if (confirm('<%=getagentPaymentLngStr("LtxtConfAddImpNoMatch")%>'.replace('{0}', '<%=txtConfVar0%>').replace('{1}', '<%=rc("DocCur")%> ' + varSaldo)))
			{
				document.Form1.finish.value = 'Y';
			}
			else
			{
				return false;
			}
		}
	}
	return true;
}


<% If Request("Err") = "importe" Then %>alert('<%=getagentPaymentLngStr("LtxtImpNotEqZero")%>'.replace('{0}', '<%=LCase(txtRct)%>'));<% End If %>

function Start(theURL, popW, popH, scroll) { // V 1.0
var winleft = (screen.width - popW) / 2;
var winUp = (screen.height - popH) / 2;
winProp = 'width='+popW+',height='+popH+',left='+winleft+',top='+winUp+',toolbar=no,scrollbars='+scroll+',menubar=no,location=no,resizable=no,status=yes'
theURL2 = 'voucher=NULL&pop=Y&AddPath=';
OpenWin = window.open('', "CtrlWindow", winProp)
doMyLink(theURL, theURL2, 'CtrlWindow');
if (parseInt(navigator.appVersion) >= 4) { OpenWin.window.focus(); }

}

</script>
<form method="POST" action="payments/submit.asp" name="Form1">
<div align="center">
<table border="0" cellpadding="0" width="100%">
	<tr class="TablasTituloSec">
		<td><%=txtRct%></td>
	</tr>
	<tr class="CanastaTitle2">
		<td><%=CmpName%></td>
	</tr>
	<tr class="CanastaTblResaltada">
		<td><p align="center"><%=getagentPaymentLngStr("DtxtLogNum")%>&nbsp;<%=Session("PayRetVal")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tr>
				<td valign="top" width="50%">
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<td class="GeneralTblBold2" width="62"><%=getagentPaymentLngStr("DtxtCode")%></td>
						<td class="GeneralTbl">
						<input readonly  type="text" class="inputDis" name="CardCode" size="40" value="<%=Replace(myHTMLEncode(RC("CardCode")), """", "&quot;")%>"></td>
					</tr>
					<tr>
						<td class="GeneralTblBold2" width="62"><%=getagentPaymentLngStr("DtxtName")%></td>
						<td class="GeneralTbl">
						<input type="text" class="input" name="CardName" size="40" value="<%=Replace(myHTMLEncode(RC("CardName")), """", "&quot;")%>" onfocus="this.select()" onkeydown="return chkMax(event, this, 100);" onchange="doProc('CardName', 'S', this.value);"></td>
					</tr>
					<tr>
						<td class="GeneralTblBold2" width="62" valign="top">
						<%=getagentPaymentLngStr("DtxtAddress")%></td>
						<td class="GeneralTbl"><textarea rows="4" class="input" name="Address" cols="31" onkeydown="return chkMax(event, this, 254);" onchange="doProc('Address', 'S', this.value);"><%=myHTMLEncode(RC("Address"))%></textarea></td>
					</tr>
					<tr>
						<td class="GeneralTblBold2" width="62"><%=getagentPaymentLngStr("LtxtTo")%></td>
						<td class="GeneralTbl"><select size="1" class="input" name="CntctCode" onchange="doProc('CntctCode', 'N', this.value);">
						<% do while not rx.eof %>
						<option <% If rc("CntctCode") = rx("CntctCode") Then response.write "selected" %> value="<%=rx("CntctCode")%>"><%=myHTMLEncode(rx("Name"))%></option>
						<% rx.movenext
						loop %>
						</select></td>
					</tr>
				<% 
				
					set rcOpt = Server.CreateObject("ADODB.RecordSet")
					cmd.CommandText = "DBOLKGetUDFWriteCols" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					cmd("@TableID") = "ORCT"
					cmd("@UserType") = "V"
					cmd("@OP") = "O"
					rcOpt.open cmd, , 3, 1

					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004

					rcOpt.Filter = "Pos = 'I'"
					do while not rcOpt.eof 
                    ShowAddPymntUFD()
                    rcOpt.movenext
                    loop 
                    If rcOpt.RecordCount > 0 then rcOpt.movefirst %>
				</table>
				</td>
				<td valign="top">
				<table border="0" cellpadding="0" width="100%" id="table4">
					<tr width="50%">
						<td class="GeneralTblBold2">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr class="GeneralTblBold2">
								<td><%=getagentPaymentLngStr("DtxtDate")%></td>
								<td width="16"><img border="0" src="images/cal.gif" id="btnDocDate"></td>
							</tr>
						</table>
						</td>
						<td class="GeneralTbl">
						<input readonly type="text" name="DocDate" class="input" size="12" value="<%=FormatDate(RC("DocDate"), False)%>" onclick="btnDocDate.click();" onchange="doProc('DocDate', 'D', this.value);"></td>
					</tr>
					<% If Not myApp.ORCTContraComp Then %>
					<tr>
						<td class="GeneralTblBold2" width="140">
						<%=getagentPaymentLngStr("LtxtCounterRef")%></td>
						<td class="GeneralTbl">
						<input type="text" name="CounterRef" class="input" size="8" value="<% If Not IsNull(RC("CounterRef")) Then %><%=myHTMLEncode(RC("CounterRef"))%><% End If %>" onkeydown="return chkMax(event, this, 8);" onchange="doProc('CounterRef', 'S', this.value);"></td>
					</tr>
					<% Else %>
					<input type="hidden" name="CounterRef" value="<% If Not IsNull(RC("CounterRef")) Then %><%=myHTMLEncode(RC("CounterRef"))%><% End If %>">
					<% End if %>
                    <% rcOpt.Filter = "Pos = 'D'"
                    do while not rcOpt.eof 
                    ShowAddPymntUFD() 
                    rcOpt.movenext
                    loop %>
					<% If ShowCurncy = "Y" Then %>
					<tr>
						<td class="GeneralTblBold2"><nobr><%=getagentPaymentLngStr("DtxtCurr")%></nobr></td>
						<td class="GeneralTbl">
						<% If EnableMC = "N" Then %><input type="text" size="3" class="inputDis" readonly name="DocCur" value="<%=myHTMLEncode(DocCur)%>"><% Else
					     %><select size="1" name="DocCur" id="DocCur" class="input" onchange="javascript:changeCur()"><%
						 
						set cmd = Server.CreateObject("ADODB.Command")
						cmd.ActiveConnection = connCommon
						cmd.CommandType = &H0004
						cmd.CommandText = "DBOLKGetCurrencies" & Session("ID")
						cmd.Parameters.Refresh()
						cmd("@LanID") = Session("LanID")
						set rctn = cmd.execute()
						do while not rctn.eof %>
						<option <% If DocCur = rctn(0) Then %>selected<% End If %> value="<%=myHTMLEncode(rctn(0))%>"><%=myHTMLEncode(rctn(1))%></option>
						<% rctn.movenext
						loop %>
						</select><% End If
						CurRate = 0
						If DocCur <> myApp.MainCur Then
							set cmd = Server.CreateObject("ADODB.Command")
							cmd.ActiveConnection = connCommon
							cmd.CommandType = &H0004
							cmd.CommandText = "DBOLKGetCartCurrRate" & Session("ID")
							cmd.Parameters.Refresh()
							cmd("@LanID") = Session("LanID")
							cmd("@LogNum") = Session("PayRetVal")
							cmd("@Pay") = "Y"
							set rctn = cmd.execute()
							CurRate = CDbl(rctn("Rate"))
						End If %><input type="text" size="12" class="inputDis" readonly name="DocCurRate" id="DocCurRate" style="text-align:right;<% If DocCur = myApp.MainCur Then %>display: none;<% End If %>" value="<%=FormatNumber(CurRate, myApp.RateDec)%>" ></td>
					</tr>
                    <% Else %>
					<input type="hidden" name="DocCur" value="<%=myHTMLEncode(DocCur)%>">
                    <% End If %>
					</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr class="GeneralTbl">
		<td>
		<p align="right"><a class="LinkGeneral" href="javascript:Start('payments/payCash.asp','476','130','no')"><%=getagentPaymentLngStr("LtxtCash")%></a> | 
		<a class="LinkGeneral" href="javascript:Start('payments/payCheck.asp','686','300','yes')"><%=getagentPaymentLngStr("LtxtCheck")%></a> | <a class="LinkGeneral" href="javascript:Start('payments/payTrans.asp','554','220','no')"><%=getagentPaymentLngStr("LtxtBankTrans")%></a> |<a class="LinkGeneral" href="javascript:Start('payments/payCred.asp',640,400,'yes')"> <%=getagentPaymentLngStr("LtxtCred")%></a> </td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<% If not rs.eof then %>
			<tr class="FirmTlt3">
				<td align="center" style="width: 15px">&nbsp;</td>
				<td align="center" style="width: 10%"><%=getagentPaymentLngStr("DtxtDoc")%></td>
				<td align="center" style="width: 10%"><%=getagentPaymentLngStr("DtxtInstallment")%></td>
				<td align="center"><%=getagentPaymentLngStr("DtxtDate")%></td>
				<td align="center"><%=getagentPaymentLngStr("DtxtDetail")%></td>
				<td align="center"><%=getagentPaymentLngStr("DtxtDocType")%></td>
				<td align="center"><%=getagentPaymentLngStr("DtxtTotal")%></td>
				<td align="center"><%=getagentPaymentLngStr("LtxtBalToPay")%></td>
				<td align="center" width="145"><%=getagentPaymentLngStr("LcolToPay")%></td>
			</tr>
			<% do while not rs.eof %>
			<tr class="GeneralTbl">
				<td style="width: 15px">
				<p align="center">
				<img src="images/checkbox_<% If CDbl(rs("SumApplied")) <> 0 Then %>on<% Else %>off<% End If %>.jpg" border="0" onclick="checkbox(this, <%=rs("DocType")%>, <%=RS("DocNum")%>, <%=rs("InstID")%>,countRow<%=RS("DocType")%>_<%=RS("DocNum")%>_<%=rs("InstID")%>,document.Form1.pay<%=RS("DocType")%>_<%=RS("DocNum")%>_<%=rs("InstID")%>,<%=getNumeric(RS("Saldo"))%>, '<%=myHTMLEncode(rs("DocCur"))%>')">
				<input type="checkbox" name="countRow<%=RS("DocType")%>_<%=RS("DocNum")%>_<%=rs("InstID")%>" id="countRow<%=RS("DocType")%>_<%=RS("DocNum")%>_<%=rs("InstID")%>" value="ON" style="display: none;" <% If CDbl(rs("SumApplied")) <> 0 Then %>checked<% end if %>></td>
				<td style="width: 10%"><a href="javascript:goDetail(<%=rs("DocType")%>, '<%=rs("DocEntry")%>')"><img id="docLink" src="design/<%=SelDes%>/images/<%=Session("rtl")%>flecha_selec.gif" border="0"></a><%=RS("DocNum")%></td>
				<td style="width: 10%">&nbsp;<%=Replace(Replace(getagentPaymentLngStr("DtxtXofY"), "{0}", RS("InstID")), "{1}", rs("InstCount"))%></td>
				<td><%=FormatDate(RS("DocDate"), True)%>&nbsp;</td>
				<td><%=RS("Comments")%>&nbsp;</td>
				<td align="center"><% Select Case rs("DocType")
					Case 13 %><%=getagentPaymentLngStr("LtxtTransTypeIN")%><%
					Case 203 %><%=getagentPaymentLngStr("LtxtTransTypeDT")%><%
					End Select %></td>
				<td align="right"><nobr><%=rs("DocCur")%>&nbsp;<%=FormatNumber(RS("DocTotal"),myApp.SumDec)%></nobr></td>
				<td align="right"><nobr><%=rs("DocCur")%>&nbsp;<%=FormatNumber(RS("Saldo"),myApp.SumDec)%></nobr></td>
				<td width="145" align="right">
				<input <% If CDbl(Rs("SumApplied")) = 0 Then %>disabled class="InputDes"<% Else %>class="input"<% end if %> type="text" name="pay<%=RS("DocType")%>_<%=RS("DocNum")%>_<%=rs("InstID")%>" id="pay" size="20" value="<%=myHTMLEncode(rs("DocCur"))%>&nbsp;<% If CDbl(Rs("SumApplied")) = 0 Then %><%=FormatNumber(RS("Saldo"),myApp.SumDec)%><% Else %><%=FormatNumber(rs("SumApplied"),myApp.SumDec)%><% End If %>" onclick="javascript:this.select()" onchange="chkThis(this, 'B', 'S', null, 0);doProcLine('SumApplied', 'N', this.value.replace(Cur<%=RS("DocType")%>_<%=RS("DocNum")%>_<%=rs("InstID")%> + ' ', ''), <%=rs("DocType")%>, <%=RS("DocNum")%>, <%=rs("InstID")%>);" style="text-align: right; width: 100%;"></td>
			</tr>
			<input type="hidden" name="Cur<%=RS("DocType")%>_<%=RS("DocNum")%>_<%=rs("InstID")%>" id="Cur<%=RS("DocType")%>_<%=RS("DocNum")%>_<%=rs("InstID")%>" value="<%=myHTMLEncode(rs("DocCur"))%>">
			<% rs.movenext
			loop %>
			<% End If %>
		</table>
		</td>
	</tr>
	<tr class="GeneralTbl">
		<td><hr size="1"></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="59%">
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<td width="90" valign="top" class="GeneralTblBold2"><%=getagentPaymentLngStr("DtxtNote")%></td>
						<td class="GeneralTbl"><textarea rows="5" name="Comments" class="input" cols="49" onkeydown="return chkMax(event, this, 254);" onchange="doProc('Comments', 'S', this.value);"><% If Not IsNull(RC("Comments")) Then %><%=myHTMLEncode(RC("Comments"))%><% End If %></textarea></td>
					</tr>
					<tr>
						<td width="90" class="GeneralTblBold2"><%=getagentPaymentLngStr("DtxtObservations")%></td>
						<td class="GeneralTbl">
						<input type="text" name="JrnlMemo" class="input" size="61" value="<% If Not IsNull(RC("JrnlMemo")) Then %><%=myHTMLEncode(RC("JrnlMemo"))%><% End If %>" onkeydown="return chkMax(event, this, 50);" onchange="doProc('JrnlMemo', 'S', this.value);"></td>
					</tr>
				</table>
				</td>
				<td valign="top" width="40%">
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<% SaldoFuera = CDbl(rc("SaldoFuera"))
						If SaldoFuera > 0 Then SaldoFuera = 0 %>
						<td class="GeneralTblBold2" style="height: 15px"><nobr><%=getagentPaymentLngStr("LtxtPayOnAcct")%></nobr></td>
						<td align="right" <% If SaldoFuera >= 0 Then %>class="GeneralTbl"<% End If %> style="height: 15px">
						<% If SaldoFuera < 0 Then %>
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
						<tr class="GeneralTbl">
							<td>
							<input <% If rc("YesSaldoFuera") = "Y" Then %>checked<% End If %> value="Y" type="checkbox" name="SaldoFuera" id="SaldoFuera" onclick="javascript:saldofuera(this, <%=getNumeric(FormatNumber(SaldoFuera*-1,myApp.SumDec))%>);" class="GeneralTbl" style="background:background-image;border: 0px solid">
							</td>
							<td align="right">
							<% End If %><label for="SaldoFuera"><nobr><%=rc("DocCur")%>&nbsp;<%=FormatNumber(SaldoFuera,myApp.SumDec)%></nobr></label><% If SaldoFuera < 0 Then %>
							</td>
							</tr>
							</table><% End If %>
						</td>
					</tr>
					<tr>
						<td class="GeneralTblBold2" style="height: 15px"><%=getagentPaymentLngStr("DtxtPaid")%></td>
						<td class="GeneralTbl" align="right" style="height: 15px">
						<input readonly type="text" name="pagado" size="18" value="<%=rc("DocCur")%>&nbsp;<%=FormatNumber(vPagado,myApp.SumDec)%>" class="GeneralTbl" style="background:background-image;border: 0px solid; text-align: right; width: 100%; "></td>
					</tr>
					<tr>
						<td class="GeneralTblBold2" style="height: 15px"><%=getagentPaymentLngStr("DtxtImport")%></td>
						<td class="GeneralTbl" style="height: 15px">
						<input readonly type="text" name="importe" size="18" value="<%=rc("DocCur")%>&nbsp;<%=FormatNumber(CDbl(rc("DocTotal")),myApp.SumDec)%>" class="GeneralTbl" style="background:background-image;border: 0px solid; text-align: right; width:100%; "></td>
					</tr>
					</table>
				</td>
			</tr>
			<tr class="GeneralTbl">
				<td colspan="2">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td>
						<input type="button" value="<% If not myAut.GetObjectProperty(24, "C") or userType = "C" Then%><%=getagentPaymentLngStr("DtxtAdd")%><%else%><%=getagentPaymentLngStr("DtxtConfirm")%><%end if%>" name="B2" onclick="if (chkSubmit()) { setPayFlow(<%=Session("PayRetVal")%>); doFlowAlert(); }"></td>
						<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
						<input type="button" class="BtnCancel" value="<%=cartBtnAddStr%><%=getagentPaymentLngStr("DtxtCancel")%>" name="I3" onclick="javascript:if(confirm('<%=getagentPaymentLngStr("LtxtConfCancelDoc")%>'))window.location.href='cartCancel.asp?RetVal=<%=Session("PayRetVal")%>'"></td>
					</tr>
				</table></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr class="GeneralTbl">
		<td>
		<p align="right"><a class="LinkGeneral" href="javascript:Start('payments/payCash.asp','476','130','no')"><%=getagentPaymentLngStr("LtxtCash")%></a> | 
		<a class="LinkGeneral" href="javascript:Start('payments/payCheck.asp','686','300','yes')"><%=getagentPaymentLngStr("LtxtCheck")%></a> | <a class="LinkGeneral" href="javascript:Start('payments/payTrans.asp','554','166','no')"><%=getagentPaymentLngStr("LtxtBankTrans")%></a> |<a class="LinkGeneral" href="javascript:Start('payments/payCred.asp',640,400,'yes')"> <%=getagentPaymentLngStr("LtxtCred")%></a> </td>
	</tr>
	<tr class="GeneralTbl">
		<td>
		<p align="right">
		&nbsp;</td>
	</tr>
	</table>
</div>
<input type="hidden" name="c1" value="<%=myHTMLEncode(rc("CardCode"))%>">
<input type="hidden" name="submitCmd" value="update">
<input type="hidden" name="Confirm" value="N">
<input type="hidden" name="DocConf" value="">
<input type="hidden" name="finish" value="N">
<input type="hidden" name="Draft" value="">
<% 
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetRates" & Session("ID")
cmd.Parameters.Refresh()
cmd("@Date") = DocDate
set rs = Server.CreateObject("ADODB.RecordSet")
set rs = cmd.execute()
do while not rs.eof %>
<input type="hidden" name="CurRate<%=rs("CurrCode")%>" value="<%=rs("Rate")%>">
<% rs.movenext
loop %>
<input type="hidden" name="AddPath" value="">
</form>
<% set rc = nothing
Sub ShowAddPymntUFD()  
AliasID = rcOpt("InsertID")
Select Case rcOpt("TypeID")
	Case "B", "N"
		ProcType = "N"
	Case "M", "A"
		ProcType = "S"
	Case "D"
		ProcType = "D"
End Select %>
        <tr class="GeneralTblBold2">
          <td width="34%" valign="top" style="padding-top: 2px;">
          <table border="0" cellpadding="0" cellspacing="0" width="100%">
          	<tr class="GeneralTblBold2">
          		<td><%=rcOpt("Descr")%><% If rcOpt("NullField") = "Y" Then %><font color="red">*</font><% End If %></td><% If (rcOpt("Query") = "Y" or rcOpt("TypeID") = "D") and IsNull(rcOpt("RTable")) Then %><td width="16">
				<img border="0" src="images/<% If rcOpt("TypeID") <> "D" Then %>flechaselec2<% Else %>cal<% End If %>.gif" id="btn<%=rcOpt("AliasID")%>" <% If rcOpt("Query") = "Y" Then %>onclick="datePicker('SmallQuery.asp?sType=Rec&FieldID=<%=rcOpt("FieldID")%>&pop=Y&addCmd=2<% If rcOpt("TypeID") = "A" Then %>&MaxSize=<%=rcOpt("SizeID")%><% End If %>',500,300,'yes', 'yes', document.Form1.U_<%=rcOpt("AliasID")%>)"<% End If %>></td><% End If %></tr></table></td>
          <td width="66%" class="GeneralTbl">
          <% If rcOpt("DropDown") = "Y" or not IsNull(rcOpt("RTable")) then 
    	set rctn = Server.CreateObject("ADODB.RecordSet")
		cmd.CommandText = "DBOLKGetUDFValues" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@TableID") = "ORCT"
		cmd("@FieldID") = rcOpt("FieldID")
		rctn.open cmd, , 3, 1 %>
		<select size="1" name="U_<%=rcOpt("AliasID")%>" class="input" style="width: 99%" onchange="doProc(this.name, '<%=ProcType%>', this.value);">
		<option></option>
		<% do while not rctn.eof %>
		<option value="<%=rctn(0)%>" <% If Not IsNull(rc(AliasID)) Then If CStr(rc(AliasID)) = CStr(rd(0)) Then Response.Write "Selected" %>><%=myHTMLEncode(rctn(1))%></option>
		<% rctn.movenext
		loop
		rctn.close %></select>
		<% ElseIf rcOpt("TypeID") = "M" and Trim(rcOpt("EditType")) = "" or rcOpt("TypeID") = "A" and rcOpt("EditType") = "?" Then %>
			<% If rcOpt("Query") = "Y" or rcOpt("TypeID") = "D" Then %>
			<table width="100%" cellspacing="0" cellpadding="0">
			  <tr>
			    <td>
			<% End If %>
			<textarea <% If rcOpt("TypeID") = "D" or rcOpt("Query") = "Y" Then %>readonly<% End If %> type="text" name="U_<%=rcOpt("AliasID")%>" class="input" onchange="chkThis(this, '<%=rcOpt("TypeID")%>', '<%=rcOpt("EditType")%>', <%=rcOpt("SizeID")%>, null);doProc(this.name, '<%=ProcType%>', this.value);" <% If rcOpt("Query") = "Y" Then %>onclick="datePicker('SmallQuery.asp?sType=Rec&FieldID=<%=rcOpt("FieldID")%>&pop=Y<% If rcOpt("TypeID") = "A" Then %>&MaxSize=<%=rcOpt("SizeID")%><% End If %>',500,300,'yes', 'yes', this)"<% End If %> rows="3" onfocus="this.select()" style="width: 100%"><% If Not IsNull(rc(AliasID)) Then %><%=myHTMLEncode(rc(AliasID))%><% End If %></textarea>
			<% If rcOpt("Query") = "Y" or rcOpt("TypeID") = "D" Then %>
				</td>
				<td width="16">
					<img border="0" src="images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="document.frmAddCard.U_<%=rcOpt("AliasID")%>.value = ''" style="cursor: hand">
				</td>
			  </tr>
			</table>
			<% End If %>
			<% ElseIf rcOpt("TypeID") = "A" and rcOpt("EditType") = "I" Then %>
				<table cellpadding="2" cellspacing="0" border="0">
					<tr>
						<td>
						<p align="center"><img src='pic.aspx?filename=<% If rc(AliasID) <> "" Then %><%=rc(AliasID)%><% Else %>n_a.gif<% End If %>&amp;MaxSize=180&amp;dbName=<%=Session("olkdb")%>' id="imgU_<%=rcOpt("AliasID")%>" border="1">
						<input type="hidden" name="U_<%=rcOpt("AliasID")%>" value="<%=rc(AliasID)%>"></td>
						<td width="16" valign="bottom"><img border="0" src="images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="javascript:document.Form1.U_<%=rcOpt("AliasID")%>.value = '';document.Form1.imgU_<%=rcOpt("AliasID")%>.src='pic.aspx?filename=n_a.gif&MaxSize=180&dbName=<%=Session("olkdb")%>';doProc('U_<%=rcOpt("AliasID")%>', '<%=ProcType%>', '');" style="cursor: hand"></td>
					</tr>
					<tr>
						<td colspan="2" height="22">
						<p align="center"><input type="button" value="<%=getagentPaymentLngStr("DtxtAddImg")%>" name="B1" onclick="javascript:getImg(document.Form1.U_<%=rcOpt("AliasID")%>, document.Form1.imgU_<%=rcOpt("AliasID")%>,180);"></td>
					</tr>
				</table>
		<% Else
		If rc(AliasID) <> "" Then 
			strVal = rc(AliasID)
			If rcOpt("TypeID") = "B" Then
        	Select Case rcOpt("EditType")
				Case "R"
					strVal = FormatNumber(CDbl(strVal),myApp.RateDec)
				Case "S"
					strVal = FormatNumber(CDbl(strVal),myApp.SumDec)
				Case "P"
					strVal = FormatNumber(CDbl(strVal),myApp.PriceDec)
				Case "Q"
					strVal = FormatNumber(CDbl(strVal),myApp.QtyDec)
				Case "%"
					strVal = FormatNumber(CDbl(strVal),myApp.PercentDec)
				Case "M"
					strVal = FormatNumber(CDbl(strVal),myApp.MeasureDec)
        	End Select
        	End If
		Else
			strVal = ""
		End If
		
		If rcOpt("TypeID") = "D" Then FldVal = FormatDate(FldVal, False)
		 %>
		<% If rcOpt("Query") = "Y" or rcOpt("TypeID") = "D" Then %><table width="100%" cellspacing="0" cellpadding="0"><tr><td><% End If %>
		<input <% If rcOpt("TypeID") = "D" or rcOpt("Query") = "Y" Then %>readonly<% End If %> type="text" name="U_<%=rcOpt("AliasID")%>" size="<% If rcOpt("TypeID") = "A" Then %>43<% Else %>12<% End If %>" class="input" onchange="chkThis(this, '<%=rcOpt("TypeID")%>', '<%=rcOpt("EditType")%>', <%=rcOpt("SizeID")%>, null);doProc(this.name, '<%=ProcType%>', this.value);" <% If rcOpt("TypeID") = "D" Then %>onclick="btn<%=rcOpt("AliasID")%>.click()"<% End If %> <% If rcOpt("Query") = "Y" Then %>onclick="datePicker('SmallQuery.asp?sType=Rec&FieldID=<%=rcOpt("FieldID")%>&pop=Y&addCmd=2<% If rcOpt("TypeID") = "A" Then %>&MaxSize=<%=rcOpt("SizeID")%><% End If %>',500,300,'yes', 'yes', this)"<% End If %> value="<% If Not IsNull(strVal) Then %><%=myHTMLEncode(strVal)%><% End If %>" <% If rcOpt("TypeID") <> "D" Then %>onfocus="this.select()"<% End If %> <% If rcOpt("TypeID") <> "D" Then %>style="width: 100%"<% End If %>>
		<% If rcOpt("Query") = "Y" or rcOpt("TypeID") = "D" Then %></td><td width="16"><img border="0" src="images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="document.Form1.U_<%=rcOpt("AliasID")%>.value = '';doProc('U_<%=rcOpt("AliasID")%>', '<%=ProcType%>', '');"></td></tr></table><% End If %><% End If %></td>
        </tr>
<% End Sub %>
<script type="text/javascript">
Calendar.setup({
    inputField     :    "DocDate",     // id of the input field
    ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
    button         :    "btnDocDate",  // trigger for the calendar (button ID)
    align          :    "Bl",           // alignment (defaults to "Bl")
    singleClick    :    true
});
<% rcOpt.Filter = "TypeID = 'D'"
If rcOpt.recordcount > 0 Then rcOpt.movefirst
do while not rcOpt.eof %>
    Calendar.setup({
        inputField     :    "U_<%=rcOpt("AliasID")%>",     // id of the input field
        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
        button         :    "btn<%=rcOpt("AliasID")%>",  // trigger for the calendar (button ID)
        align          :    "Bl",           // alignment (defaults to "Bl")
        singleClick    :    true
    });
<% rcOpt.movenext
loop 
set rcOpt = Nothing %>


</script>
<form target="_blank" method="post" name="frmViewDetail" action="">
<input type="hidden" name="DocEntry" value="">
<input type="hidden" name="DocType" value="">
<input type="hidden" name="CardCode" value="">
<input type="hidden" name="pop" value="Y">
</form>
<script language="javascript">
function goDetail(DocType, DocEntry) {
	if (DocType == 2)
	{
		document.frmViewDetail.action = 'addCard/crdConfDetailOpen.asp';
		document.frmViewDetail.CardCode.value = DocEntry;
	}
	else if (DocType != 24)
	{
		document.frmViewDetail.action = "cxcDocDetailOpen.asp";
		document.frmViewDetail.DocEntry.value = DocEntry;
	}
	document.frmViewDetail.DocType.value = DocType;
	document.frmViewDetail.submit();
}
</script>

<!--#include file="agentBottom.asp"-->