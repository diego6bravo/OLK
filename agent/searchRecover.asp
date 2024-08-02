<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Session("useraccess") <> "P" Then Response.Redirect "unauthorized.asp" %>
<% addLngPathStr = "" %>
<!--#include file="lang/searchRecover.asp" -->

<%
If Request("btnRecover") <> "" Then
	sql = ""
	ArrVal = Split(Request("rLogNum"),", ")
	for i = 0 to UBound(ArrVal)
		sql = sql & "update R3_ObsCommon..tlog set Status = 'R' where LogNum = " & ArrVal(i) & " "
	next
	if sql <> "" then conn.execute(sql)
ElseIf Request("btnRemove") <> "" Then
	ArrVal = Split(Request("rLogNum"),", ")
	for i = 0 to UBound(ArrVal)
		sql = sql & "update R3_ObsCommon..tlog set Status = 'B' where LogNum = " & ArrVal(i) & " "		
	next
	if sql <> "" then conn.execute(sql)
End If
           AsignedSlp =  not myAut.HasAuthorization(60)

LogNum = ""
If Request("LogNumFrom") <> "" Then LogNum = " and T0.LogNum >= " & Request("LogNumFrom") & " "
If Request("LogNumTo") <> "" Then LogNum = LogNum & " and T0.LogNum <= " & Request("LogNumTo") & " "

DocTypeAdd = ""
If Request("dtFrom") <> "" Then DocTypeAdd = " and DateDiff(day,Convert(datetime,'" & SaveSqlDate(Request("dtFrom")) & "',120), DocDate) >= 0 "
If Request("dtTo") <> "" Then DocTypeAdd = DocTypeAdd & " and DateDiff(day,DocDate, Convert(datetime,'" & SaveSqlDate(Request("dtTo")) & "',120)) >= 0"

cCode = ""
If Request("CardCodeFrom") <> "" Then cCode = " and C1.CardCode >= N'" & Request("CardCodeFrom") & "' "
If Request("CardCodeTo") <> "" Then cCode = cCode & " and C1.CardCode <= N'" & Request("CardCodeTo") & "' "

GroupCode = ""
If Request("GroupNameFrom") <> "" Then GroupCode = " and C3.GroupName >= N'" & Request("GroupNameFrom") & "' "
If Request("GroupNameTo") <> "" Then GroupCode = GroupCode & " and C3.GroupName <= N'" & Request("GroupNameTo") & "' "

Country = ""
If Request("CountryFrom") <> "" Then Country = " and C2.Name >= N'" & Request("CountryFrom") & "' "
If Request("CountryTo") <> "" Then Country = Country & " and C2.Name <= N'" & Request("CountryTo") & "' "

If Request("DocType") <> "" Then 
	DocType = "and Object = " & Replace(Request("DocType"),"C","")
	If Request("DocType") = "13" Then DocTypeAdd = " and Not Exists (select 'A' from olkcic where invlognum = T0.LogNum) "
	If Request("DocType") = "13C" Then DocTypeAdd = " and Exists (select 'A' from olkcic where invlognum = T0.LogNum) "
End If

If Request("orden1") <> "" Then 
	orden1 = Request("orden1") & " "
	orden2 = Request("orden2")
Else
	orden1 = "1 desc"
End If
If AsignedSlp Then 
	SlpCode1 = "and T1.SlpCode = @SlpCode "
	SlpCode2 = "and T2.SlpCode = @SlpCode "
End If

strStatus = ""
If Request("SearchDel") = "" Then
	strStatus = "status = 'E'"
Else
	strStatus = "status in ('E', 'B')"
End If

sqll = _
"Declare @Company nvarchar(20) set @Company = N'" & Session("olkdb") & "'" & _
"Declare @SlpCode int set @SlpCode = " & Session("VendId") & " " & _
"select T0.LogNum, IsNull(SlpName, '') SlpName, T1.CardCode, IsNull(T1.CardName, '') CardName, IsNull(name, '') Country, IsNull(GroupName, '') GroupName, " & _
"IsNull(Comments, '') Comments, DocDate, Object, Status, DocCur Currency, " & _
"OLKCommon.dbo.DBOLKDocTotal" & Session("ID") & "(T0.LogNum) DocTotal, " & _
"case when exists(select 'A' from ocrd where cardcode = T1.cardcode collate database_default) Then 'True' ELSE 'False' End Verfy, " & _
"(select PayLogNum from olkCIC where InvLogNum = T0.LogNum) PayLogNum, DocNum = (select DocNum from olkdoceditcontrol where lognum = T0.LogNum) " & _
"from R3_Obscommon..tlog T0 inner join r3_obscommon..tdoc T1 on T1.LogNum = T0.LogNum " & _
"inner join ocrd C1 on C1.CardCode = T1.CardCode collate database_default " & _
"left outer join ocry C2 on C2.code = C1.Country " & _
"inner join ocrg C3 on C3.GroupCode = C1.GroupCode " & _
"inner join oslp on oslp.slpcode = T1.SlpCode " & _
"inner join R3_ObsCommon..TLOGControl T2 on T2.logNum = T0.LogNum and T2.AppID = 'TM-OLK' " & _
"where Company = @Company and Object in (17,23,15,13) and " & strStatus & " " & SlpCode1 & _
" and T1.SlpCode <> (select slpcode from olkcommon) " & _
cCode & LogNum & DocDate & CardString & Country & GroupCode & DocType & DocTypeAdd & _
"union select T0.LogNum, IsNull(SlpName, '') SlpName, T1.CardCode, IsNull(T1.CardName, '') CardName, IsNull(name, '') Country, IsNull(GroupName, '') GroupName, " & _
"IsNull(Comments, '') Comments, DocDate, Object, Status, DocCur Currency, " & _
"Case T2.DocType When 13 Then (select ISNULL(sum(SumApplied),0) from r3_obscommon..pmt2 where lognum = T0.Lognum) " & _ 
"When 17 Then (select ISNULL(sum(SumApplied),0) from OLKRCP where PayLogNum = T0.Lognum) End DocTotal, " & _
"case when exists(select 'A' from ocrd where cardcode = T1.cardcode collate database_default) Then 'True' ELSE 'False' End Verfy, PayLogNum = NULL, DocNum = NULL " & _
"from r3_obscommon..tlog T0 " & _
"inner join r3_obscommon..tpmt T1 on T1.LogNum = T0.LogNum " & _
"inner join OLKDocControl T2 on T2.LogNum = T0.LogNum " & _
"inner join ocrd C1 on C1.CardCode = T1.CardCode collate database_default " & _
"left outer join ocry C2 on C2.code = C1.Country " & _
"inner join ocrg C3 on C3.GroupCode = C1.GroupCode " & _
"inner join oslp on oslp.slpcode = T2.SlpCode " & _
"inner join R3_ObsCommon..TLOGControl T3 on T3.logNum = T0.LogNum and T3.AppID = 'TM-OLK' " & _
"where Company = @Company and Object = (24) and " & strStatus & " " & _
"and not exists(select 'A' from olkcic where PayLogNum = T0.LogNum) " & _
SlpCode2 & cCode & LogNum & DocDate & CardString & Country & GroupCode & DocType & _
"order by " & orden1 & orden2
'response.write sqll

set rs = Server.CreateObject("ADODB.RecordSet")
RS.CursorLocation = 3 ' adUseClient
rs.open sqll, conn
RS.PageSize = 40
nPageCount = RS.PageCount
nPage = CLng(Request("Page"))
If nPage < 1 Or nPage > nPageCount Then	nPage = 1
If Not Rs.Eof then RS.AbsolutePage = nPage
%>
<script language="javascript">
<!--
function docAlert(dType) {
var DocType;
switch (dType) {
	case 17:
		DocType = "<%=txtOrdr%>";
		break
	case 23:
		DocType = "<%=txtQuote%>";
		break
	case 24:
		DocType = "<%=txtRct%>";
		break
	case 48:
		DocType = "<%=txtInv%>/<%=txtRct%>";
		break
	case 13:
		DocType = "<%=txtInv%>";
		break
	case 15:
		DocType = "<%=txtOdln%>";
		break
}
alert('<%=getsearchRecoverLngStr("LvalDisDocs")%>'.replace('{0}', DocType));
}
//-->
function valFrm(frm)
{
var LogNums = frm.rLogNum;
var RetVal = false;
<% if rs.recordcount > 1 then %>
	for (var i = 0;i < LogNums.length;i++)
	{
		if (LogNums(i).checked == true) { RetVal = true; }
	}
<% Else %>
	if (LogNums.checked == true) { RetVal = true; }
<% End If %>
if (!RetVal) { alert('<%=getsearchRecoverLngStr("LtxtValNoDoc")%>'); }
return RetVal;
}
function goP(p) { document.frmRec.page.value = p; document.frmRec.submit(); }
</script>
<body style="text-align: right">

<div align="center">
<form method="POST" action="searchRecover.asp" name="frmRec" onsubmit="return valFrm(this);">
<table border="0" cellpadding="0" width="100%" id="table1">
	<tr class="GeneralTlt">
		<td colspan="3"><%=getsearchRecoverLngStr("LttlRecDocs")%></td>
	</tr>
	<tr class="GeneralTbl">
		<td>
		<p align="center">
		<% If nPage > 1 Then %><a href='#' onclick="goP(<%= nPage - 1 %>);"><img border="0" src="design/0/images/<%=Session("rtl")%>prev_icon_trans.gif" width="15" height="15"></a><% End If %></td>
		<td><p align="center"><% if rs.PageCount > 1 then
               For I = 1 To rs.PageCount
				If I = nPage Then %>
				<b><font size="1"><%= I %></font></b>
				<% Else %>
				<a href='#' onclick="goP(<%= I %>);"><%= I %></a>
				<% End If
				Next 'I
				end if %></td>
		<td>
		<p align="center">
		<% If nPage < rs.PageCount Then %><a href='#' onclick="goP(<%= nPage + 1 %>);"><img border="0" src="design/0/images/<%=Session("rtl")%>next_icon_trans.gif" width="15" height="15"></a><% End If %></td>
	</tr>
	<tr>
		<td colspan="3">
		<table border="0" cellpadding="0" width="100%" id="table2">
			<tr class="GeneralTblBold2">
				<td align="center">&nbsp;</td>
				<td align="center">&nbsp;</td>
				<td align="center"><%=getsearchRecoverLngStr("DtxtLogNum")%></td>
				<td align="center"><% If 1 = 2 Then %><%=getsearchRecoverLngStr("DtxtAgent")%><% Else %><%=txtAgent%><% End If %></td>
				<td align="center"><% If 1 = 2 Then %><%=getsearchRecoverLngStr("DtxtClients")%><% Else %><%=txtClient%><% End If %></td>
				<td align="center"><%=getsearchRecoverLngStr("DtxtName")%></td>
				<td align="center"><%=getsearchRecoverLngStr("DtxtGroup")%></td>
				<td align="center"><%=getsearchRecoverLngStr("DtxtCountry")%></td>
				<td align="center"><%=getsearchRecoverLngStr("DtxtDate")%></td>
				<td align="center"><%=getsearchRecoverLngStr("DtxtType")%></td>
				<td align="center"><%=getsearchRecoverLngStr("DtxtTotal")%></td>
			</tr>
		  <%  if not rs.eof then
		  do while not (rs.eof Or RS.AbsolutePage <> nPage )
		  Enable = True
		  Select Case rs("Object")
		  	Case 13
		  		If IsNULL(rs("PayLogNum")) Then
		  			If Not myApp.EnableOINV Then Enable = False
		  		Else
		  			If Not myApp.EnableCashInv Then Enable = False
		  		End If
		  	Case 17
		 			If Not myApp.EnableORDR Then Enable = False
		  	Case 15
		 			If Not myApp.EnableODLN Then Enable = False
		  	Case 23
		  			If Not myApp.EnableOQUT Then Enable = False
		  	Case 24
		  			If Not myApp.EnableORCT Then Enable = False
		  	Case 203
		  			If Not myApp.EnableODPIReq Then Enable = False
		  	Case 204
		  			If Not myApp.EnableODPIInv Then Enable = False
		  End Select %>
			<tr class="GeneralTbl">
				<td>
				<input type="checkbox" name="rLogNum" value="<%=rs("LogNum")%>" style="border-style:solid; border-width:0; background:background-image"></td>
				<td>
				<p align="center">
				<a href="javascript:GoLogView('<% If rs("Object") <> "24" Then %>cxcdocdetail<% Else %>cxcRctDetail<% End If %>Open.asp',<%=rs("LogNUm")%>)">
				<img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></a></td>
				<td><%=RS("LogNum")%>&nbsp;</td>
				<td><%=RS("SlpName")%>&nbsp;</td>
				<td><%=RS("CardCode")%>&nbsp;</td>
				<td><%=RS("cardname")%>&nbsp;</td>
				<td><%=RS("GroupName")%>&nbsp;</td>
				<td>
				<%=RS("Country")%>&nbsp;</td>
				<td><%=FormatDate(RS("DocDate"), True)%>&nbsp;</td>
				<td><%
			     If rs("DocNum") <> "" then response.write getsearchQryLngStr("LtxtEditOf") & "&nbsp;"
			     Select Case RS("Object")
			     Case 17
			     Response.write txtOrdr '"Pedido"
			     Case 23
			     Response.write txtQuote '"Cotización"
			     Case 24
			     Response.write txtRct '"Recibo"
			     Case 13
			     Response.write txtInv '"Factura"
			     Case 14
			     Response.write txtOrin '"Nota de credito"
			     Case 16
			     Response.write txtOrdn '"Devoluciones"
			     Case 18
			     Response.write txtOpch '"Comprobante de compra"
			     Case 20
			     Response.write "Consignación de mercancia"
			     Case 59
			     Response.write txtOpdn '"Entreda general al inventario"
			     Case 15
			     Response.write txtOdln '"Despachos"
			     Case 19
			     Response.write txtOrpc '"Nota de debito"
			     Case 21
			     Response.write "Devoluciones en compra"
			     Case 60
			     Response.write "Salida general del inventario"
			     Case 67
			     Response.write "Transferencia entre bodegas"
			     End Select 
			     If rs("DocNum") <> "" then response.write " #" & rs("DocNum")
			     If rs("PayLogNum") <> "" Then Response.Write "/" & txtRct%></td>
				<td align="center">
				<p align="center">
				<% If Rs("Doctotal") <> 0 Then %><nobr><%=rs("Currency")%>&nbsp;<%=FormatNumber(rs("DocTotal"),myApp.PriceDec)%></nobr><% end if %></td>
			</tr>
			<% If rs("Comments") <> "" Then %>
			<tr class="GeneralTbl">
				<td>
				&nbsp;</td>
				<td>
				&nbsp;</td>
				<td colspan="9"><%=getsearchRecoverLngStr("DtxtObservations")%>: <%=rs("Comments")%></td>
			</tr>
			<% End If %>
		 <% rs.movenext
		 loop %>
			<tr class="GeneralTbl">
				<td colspan="11">
				<table border="0" cellpadding="0" width="100%" id="table3">
					<tr class="GeneralTbl">
						<td width="349">
						<input type="submit" value="<%=getsearchRecoverLngStr("LtxtRecover")%>" name="btnRecover"></td>
						<td>
						<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
						<input type="submit" value="<%=getsearchRecoverLngStr("LtxtDel")%>" name="btnRemove" onclick="return confirm('<%=getsearchRecoverLngStr("LtxtConfDel")%>')"></td>
					</tr>
				</table>
				</td>
			</tr>
			<% Else %>
			<tr class="GeneralTblBold2">
				<td colspan="11">
				<p align="center"><%=getsearchRecoverLngStr("LtxtNoRecDocs")%></td>
			</tr>
			<% End If %>
		</table>
		</td>
	</tr>
	<tr class="GeneralTbl">
		<td>
		<p align="center">
		<% If nPage > 1 Then %><a href='#' onclick="goP(<%= nPage - 1 %>);"><img border="0" src="design/0/images/<%=Session("rtl")%>prev_icon_trans.gif" width="15" height="15"></a><% End If %></td>
		<td><p align="center"><% if rs.PageCount > 1 then
               For I = 1 To rs.PageCount
				If I = nPage Then %>
				<b><font size="1"><%= I %></font></b>
				<% Else %>
				<a href='#' onclick="goP(<%= I %>);"><%= I %></a>
				<% End If
				Next 'I
				end if %></td>
		<td>
		<p align="center">
		<% If nPage < rs.PageCount Then %><a href='#' onclick="goP(<%= nPage + 1 %>);"><img border="0" src="design/0/images/<%=Session("rtl")%>next_icon_trans.gif" width="15" height="15"></a><% End If %></td>
	</tr>
</table>
<% for each item in Request.Form
if item <> "rLogNum" and item <> "page" then %>
<input type="hidden" name="<%=item%>" value="<%=Request.Form(item)%>">
<% end if
Next %>
<% for each item in Request.QueryString 
if item <> "rLogNum" then %>
<input type="hidden" name="<%=item%>" value="<%=Request.QueryString(item)%>">
<% end if
Next %>
<input type="hidden" name="page" value="<%=nPage%>">
</form>
    </div>
<script language="javascript">
function GoLogView(Action, LogNum) {
document.viewLogNum.action = Action
document.viewLogNum.DocEntry.value = LogNum 
document.viewLogNum.submit() }
</script>
<form target="_blank" method="post" name="viewLogNum">
<input type="hidden" name="DocEntry" value="">
<input type="hidden" name="DocType" value="-2"></form>
<!--#include file="agentBottom.asp"-->