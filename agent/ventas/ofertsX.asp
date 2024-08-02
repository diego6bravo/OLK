<% addLngPathStr = "ventas/" %>
<!--#include file="lang/ofertsX.asp" -->
<%
varxx = 0
varx = 0
If Request("Orden1") <> "" Then
orden1 = Request("Orden1")
orden2 = Request("Orden2")
ElseIf Request("Orden1") = "" Then
orden1 = "8 "
orden2 = "desc "
End If

iPageSize = 10
set rs = Server.CreateObject("ADODB.recordset")
SaleType = myApp.GetSaleUnit

           
If Request("page") = "" Then
iPageCurrent = 1

sqlFilter = ""
If Request("dtBy") = "O" Then varBy = "ofert" Else varBy = "response"
If Request("dtFrom") <> "" Then sqlFilter = " and DateDiff(day, " & varBy & "Date, Convert(datetime,'" & SaveSqlDate(Request("dtFrom")) & "',120)) <= 0"
If Request("dtTo") <> "" Then sqlFilter = sqlFilter & " and DateDiff(day, Convert(datetime,'" & SaveSqlDate(Request("dtTo")) & "',120), " & varBy & "Date) <= 0"
If Request("CardCodeFrom") <> "" Then sqlFilter = sqlFilter & " and T0.UserName >= N'" & saveHTMLDecode(Request("CardCodeFrom"), False) & "' "
If Request("CardCodeTo") <> "" Then sqlFilter = sqlFilter & " and T0.UserName <= N'" & saveHTMLDecode(Request("CardCodeTo"), False) & "' "
If strScriptName = "activeclient.asp" Then sqlFilter = sqlFilter & " and T0.UserName = N'" & saveHTMLDecode(Session("UserName"), False) & "' "
If Request("ItemCodeFrom") <> "" Then sqlFilter = sqlFilter & " and T0.ItemCode >= N'" & saveHTMLDecode(Request("ItemCodeFrom"), False) & "' "
If Request("ItemCodeTo") <> "" Then sqlFilter = sqlFilter & " and T0.ItemCode <= N'" & saveHTMLDecode(Request("ItemCodeTo"), False) & "' "
If Request("PriceBy") = "O" Then varBy = "ofert" Else varBy = "response"
If Request("PriceFrom") <> "" Then sqlFilter = sqlFilter & " and " & varBy & "IsNull(Price, 0) >= " & Request("PriceFrom") & " "
If Request("PriceTo") <> "" Then sqlFilter = sqlFilter & " and " & varBy & "IsNull(Price, 0) <= " & Request("PriceTo") & " "
If Request("QtyBy") = "O" Then varBy = "ofert" Else varBy = "response"
If Request("QtyFrom") <> "" Then sqlFilter = sqlFilter & " and " & varBy & "Quantity >= " & Request("QtyFrom") & " "
If Request("QtyTo") <> "" Then sqlFilter = sqlFilter & " and " & varBy & "Quantity <= " & Request("QtyTo") & " "
If Request("GroupNameFrom") <> "" or Request("GroupNameTo") <> "" Then
	sqlFilter = sqlFilter & " and exists(select 'A' from OCRG where GroupCode = ocrd.GroupCode "
	If Request("GroupNameFrom") <> "" Then sqlFilter = sqlFilter & " and GroupName >= N'" & saveHTMLDecode(Request("GroupNameFrom"), False) & "' "
	If Request("GroupNameTo") <> "" Then sqlFilter = sqlFilter & " and GroupName <= N'" & saveHTMLDecode(Request("GroupNameTo"), False) & "' "
	sqlFilter = sqlFilter & ")"
End If
If Request("CountryFrom") <> "" or Request("CountryTo") <> "" Then
	sqlFilter = sqlFilter & " and exists(select 'A' from OCRY where Code = ocrd.Country "
	If Request("CountryFrom") <> "" Then sqlFilter = sqlFilter & " and Name >= N'" & saveHTMLDecode(Request("CountryFrom"), False) & "' "
	If Request("CountryTo") <> "" Then sqlFilter = sqlFilter & " and Name <= N'" & saveHTMLDecode(Request("CountryTo"), False) & "' "
	sqlFilter = sqlFilter & ")"
End If
If Request("ItmsGrpNamFrom") <> "" or Request("ItmsGrpNamTo") <> "" Then
	sqlFilter = sqlFilter & " and exists(select 'A' from OITB where ItmsGrpCod = oitm.ItmsGrpCod "
	If Request("ItmsGrpNamFrom") <> "" Then sqlFilter = sqlFilter & " and ItmsGrpNam >= N'" & saveHTMLDecode(Request("ItmsGrpNamFrom"), False) & "' "
	If Request("ItmsGrpNamTo") <> "" Then sqlFilter = sqlFilter & " and ItmsGrpNam <= N'" & saveHTMLDecode(Request("ItmsGrpNamTo"), False) & "' "
	sqlFilter = sqlFilter & ")"
End If
If Request("FirmNameFrom") <> "" or Request("FirmNameTo") <> "" Then
	sqlFilter = sqlFilter & " and exists(select 'A' from OMRC where FirmCode = oitm.FirmCode "
	If Request("FirmNameFrom") <> "" Then sqlFilter = sqlFilter & " and FirmName >= N'" & saveHTMLDecode(Request("FirmNameFrom"), False) & "' "
	If Request("FirmNameTo") <> "" Then sqlFilter = sqlFilter & " and FirmName <= N'" & saveHTMLDecode(Request("FirmNameTo"), False) & "' "
	sqlFilter = sqlFilter & ")"
End If
If Request("Note") <> "" Then
	If Request("NoteBy") = "O" Then varBy = "ofert" Else varBy = "response"
	sqlFilter = sqlFilter & " and " & varBy & "Note like '%" & saveHTMLDecode(Request("Note"), False) & "%' "
End If
If Request("DueFrom") <> "" or Request("DueTo") <> "" Then
	If Request("DueBy") = "O" Then varBy = "ofert" Else varBy = "response"
	If Request("DueFrom") <> "" Then sqlFilter = sqlFilter & " and DateDiff(day,getdate(),DateAdd(day,"&varBy&"Limit, "&varBy&"Date)) >= " & Request("DueFrom")
	If Request("DueTo") <> "" Then sqlFilter = sqlFilter & " and DateDiff(day,getdate(),DateAdd(day,"&varBy&"Limit, "&varBy&"Date)) <= " & Request("DueTo")
End If
If Request("OfertStatus") <> "" Then
	ArrVal = Split(Request("OfertStatus"), ", ")
	OfertStatus = ""
	For i = 0 to UBound(ArrVal)
		OfertStatus = OfertStatus & ", '" & ArrVal(i) & "'"
	Next
	OfertStatus = Right(OfertStatus, Len(OfertStatus)-1)
	sqlFilter = sqlFilter & " and OfertStatus in (" & OfertStatus & ") "
End If

If Request("SlpCodeFrom") <> "" Then sqlFilter = sqlFilter & " and OSLP.SlpName >= N'" & saveHTMLDecode(Request("SlpCodeFrom"), False) & "' "
If Request("SlpCodeTo") <> "" Then sqlFilter = sqlFilter & " and OSLP.SlpName <= N'" & saveHTMLDecode(Request("SlpCodeTo"), False) & "' "


If Not IsNull(myApp.AgentClientsFilter) and not IgnoreGeneralFilter Then
	sqlFilter = sqlFilter & " and T0.UserName not in (" & Replace(Replace(myApp.AgentClientsFilter, "@SlpCode", Session("vendid")), "@Type", 1) & ") "
End If

If not myAut.HasAuthorization(60) Then sqlFilter = sqlFilter & " and OCRD.SLPCode = " & Session("vendid") & " "


sql = "Declare @CostPrcLst int set @CostPrcLst = (select top 1 CostPrcLst from oadm order by CurrPeriod desc) " & _
"Declare @GrossBySal char(1) set @GrossBySal = (select top 1 GrossBySal from oadm order by CurrPeriod desc) " & _
"select *, Case When CostPrice <> 0 Then " & _
"Case ofertPrice when 0 then 0 else Case @GrossBySal When 'Y' then 100-(CostPrice/ofertPrice*100) When 'N' Then ((ofertPrice/CostPrice)-1)*100 End End Else 0 End " & _
"As OfertCostPercentage,  " & _
"Case When CostPrice <> 0 Then Case responsePrice when 0 then 0 else Case @GrossBySal When 'Y' then 100-(CostPrice/responsePrice*100) When 'N' " & _
"Then ((responsePrice/CostPrice)-1)*100 End End Else 0 End As ResponseCostPercentage " & _
"from (select T0.ofertIndex, ocrd.CardCode, IsNull(CardName,ocrd.CardCode) CardName, T0.ItemCode, " & _
"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', T0.ItemCode, ItemName) ItemName, BasePrice, ofertStatus, " & _
"ofertDate, DateAdd(day,OfertLimit, ofertDate) ofertDueDate, ofertQuantity, ofertPrice, ofertDiscount, " & _
"ResponseDate, DateAdd(Day,responseLimit, ResponseDate) ResponseDueDate, ResponseQuantity, ResponsePrice, ResponseDiscount, " & _
"IsNull(SalUnitMsr, '') SalUnitMsr, NumInSale, SalPackUn, IsNull(SalPackMsr, '') SalPackMsr, " & _
"Case @CostPrcLst When -1 Then lastpurprc When -2 then  lstevlpric else " & _
"IsNull((select price from itm1 where itemcode = oitm.itemcode and pricelist = @CostPrcLst), 0) end As CostPrice, " & _
"Case @CostPrcLst When -1 Then LastPurCur When -2 then (select top 1 MainCurncy from oadm order by CurrPeriod desc) Else " & _
"(select Currency from itm1 where itemcode = oitm.itemcode and pricelist = @CostPrcLst) End As Currency  " & _
"from olkoferts T0 " & _
"inner join olkofertslines T1 on T1.ofertIndex = T0.ofertIndex " & _
"inner join ocrd on ocrd.cardcode = T0.UserName " & _
"inner join oitm on oitm.itemcode = T0.ItemCode " & _
"inner join OSLP on OSLP.SlpCode = OCRD.SlpCode " & _
"where ofertLineNum = (select max(ofertLineNum) from olkofertslines where ofertIndex = T0.ofertIndex)" & _
sqlFilter & " and transStatus = 'O') As Table1 " & _
"order by " & orden1 & orden2
'response.write sql
Session("sqlstmt") = sql
Else
	iPageCurrent = CInt(Request("page"))
	sql = Session("sqlstmt")
End If
rs.CursorType = 3
rs.CursorLocation = 3
set rs.ActiveConnection = conn
rs.open sql
           RS.PageSize = iPageSize
           RS.CacheSize = iPageSize
iPageCount = RS.PageCount
If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
If iPageCurrent < 1 Then iPageCurrent = 1
if not rs.eof then RS.AbsolutePage = iPageCurrent
%>
<SCRIPT LANGUAGE="JavaScript">
function Start(page, w, h) {
OpenWin = this.open(page, "ofertsX", "toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no,width="+w+",height="+h);
OpenWin.focus()
}
</SCRIPT>
    <div align="center">
<table border="0" cellpadding="0" width="100%" id="table1">
	<tr class="GeneralTlt">
		<td><% If 1 = 2 Then %>Manejo de ofertas<% Else %><%=Replace(getofertsXLngStr("LttlOfertMan"), "{0}", txtOferts)%><% End If %></td>
	</tr>
	<% If iPageCount > 1 Then %>
	<tr>
		<td><% doOfertsXPages %></td>
	</tr>
	<% End If %>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table2">
			<tr class="FirmTlt3" align="center">
				<td colspan="2"><% If 1 = 2 Then %><%=getofertsXLngStr("DtxtClient")%><% Else %><%=txtClient%><% End If %></td>
				<td colspan="2"><%=getofertsXLngStr("DtxtDescription")%>&nbsp;(<%=getofertsXLngStr("DtxtCode")%>)</td>
				<td><%=getofertsXLngStr("LtxtBasePrice")%></td>
				<td><%=getofertsXLngStr("LtxtCostPrice")%></td>
				<td><%=getofertsXLngStr("DtxtState")%></td>
			</tr>
			<tr class="FirmTlt3">
				<td align="center" colspan="2">&nbsp;</td>
				<td align="center"><%=getofertsXLngStr("DtxtDate")%></td>
				<td align="center"><%=getofertsXLngStr("LtxtDueDate")%></td>
				<td align="center"><%=getofertsXLngStr("DtxtQty")%></td>
				<td align="center"><%=getofertsXLngStr("DtxtPrice")%></td>
				<td align="center"><%=getofertsXLngStr("LtxtGrosProfit")%></td>
			</tr>
		  <% if not rs.eof then
		  for intRecord=1 to rs.PageSize 
				  Select Case SaleType
				  	Case 1
				  		BasePrice = CDbl(rs("BasePrice"))
				  		OfertPrice = CDbl(rs("OfertPrice"))
				  		UnPrice = "Un."
				  		SaleUn = "Un.(1)"
				  		ofertQuantity = rs("ofertQuantity")
				  		responseQuantity = rs("responseQuantity")
				  		ofertTotal = CDbl(rs("ofertQuantity"))*CDbl(rs("ofertPrice"))
				  		responseTotal = CDbl(rs("responseQuantity"))*CDbl(rs("responsePrice"))
				  		responsePrice = CDbl(rs("responsePrice"))
				  		CostPrice = CDbl(rs("CostPrice"))
				  	Case 2
				  		BasePrice = CDbl(rs("BasePrice"))*CDbl(rs("NumInSale"))
				  		OfertPrice = CDbl(rs("OfertPrice"))*CDbl(rs("NumInSale"))
				  		UnPrice = rs("SalUnitMsr")
				  		SaleUn = rs("SalUnitMsr")
				  		If myApp.GetShowQtyInUn Then SaleUn = SaleUn & "(" & rs("NumInSale") & ")"
				  		ofertQuantity = rs("ofertQuantity")
				  		responseQuantity = rs("responseQuantity")
				  		ofertTotal = CDbl(rs("ofertQuantity"))*CDbl(rs("ofertPrice"))*CDbl(rs("NumInSale"))
				  		responseTotal = CDbl(rs("responseQuantity"))*CDbl(rs("responsePrice"))*CDbl(rs("NumInSale"))
				  		responsePrice = CDbl(rs("responsePrice"))*CDbl(rs("NumInSale"))
				  		CostPrice = CDbl(rs("CostPrice"))*CDbl(rs("NumInSale"))
				  	Case 3
					  	SaleUn = rs("SalPackMsr")
					  	If myApp.GetShowQtyInUn Then SaleUn = SaleUn & "(" & rs("SalPackUn") & ")" 
				  		If myApp.UnEmbPriceSet Then
				  			BasePrice = CDbl(rs("BasePrice"))*CDbl(rs("NumInSale"))
					  		OfertPrice = CDbl(rs("OfertPrice"))*CDbl(rs("NumInSale"))
					  		responsePrice = CDbl(rs("responsePrice"))*CDbl(rs("NumInSale"))
				  			UnPrice = rs("SalUnitMsr")
					  		CostPrice = CDbl(rs("CostPrice"))*CDbl(rs("NumInSale"))
					  		SaleUn = SaleUn & " x " & rs("SalUnitMsr") 
					  		If myApp.GetShowQtyInUn Then SaleUn = SaleUn & "(" & rs("NumInSale") & ")"
				  		Else
				  			BasePrice = CDbl(rs("BasePrice"))*CDbl(rs("NumInSale"))*CDbl(rs("SalPackUn"))
					  		OfertPrice = CDbl(rs("OfertPrice"))*CDbl(rs("NumInSale"))*CDbl(rs("SalPackUn"))
					  		responsePrice = CDbl(rs("responsePrice"))*CDbl(rs("NumInSale"))*CDbl(rs("SalPackUn"))
				  			UnPrice = rs("SalPackMsr")
					  		CostPrice = CDbl(rs("CostPrice"))*CDbl(rs("NumInSale"))*CDbl(rs("SalPackUn"))
				  		End If
				  		ofertQuantity = CDbl(rs("ofertQuantity"))/CDbl(rs("SalPackUn"))
				  		responseQuantity = CDbl(rs("responseQuantity"))/CDbl(rs("SalPackUn"))
				  		ofertTotal = CDbl(rs("ofertQuantity"))*CDbl(rs("ofertPrice"))*CDbl(rs("NumInSale"))
				  		responseTotal = CDbl(rs("responseQuantity"))*CDbl(rs("responsePrice"))*CDbl(rs("NumInSale"))
				  End Select
		  %>
			<tr class="GeneralTbl">
				<td>
				<p align="center">
				<a href="javascript:GoLogView('<%=myHTMLEncode(rs("CardCode"))%>')">
				<img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></a></td>
				<td><%=rs("CardName")%>&nbsp;</td>
				<td colspan="2"><a href="javascript:goViewItem('<%=Replace(rs("ItemCode"), "'", "\'")%>')"><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></a><%=rs("ItemName")%>&nbsp;(<%=rs("ItemCode")%>)</td>
				<td>
				<p align="right"><nobr><%=rs("Currency")%>&nbsp;<%=FormatNumber(BasePrice,myApp.PriceDec)%></nobr></td>
				<td>
				<p align="right"><nobr><%=rs("Currency")%>&nbsp;<%=FormatNumber(CostPrice,myApp.PriceDec)%></nobr></td>
				<td align="center">
				<% Select Case rs("ofertStatus") 
                  Case "W"
                  	Response.write "<blink><font color=""#FF9933"">" & getofertsXLngStr("DtxtWaiting") & "</font></blink>"
                  Case "A" 
                 	Response.write "<font color=""#008080"">" & getofertsXLngStr("DtxtAproved") & "</font>"
                  Case "O"
                  	'Response.write "<font color=""#3366CC"">" & getofertsXLngStr("DtxtCounter") & " " & txtOfert & "</font>"
                  	Response.write "<font color=""#3366CC"">" & Replace(Replace(getofertsXLngStr("LbtnCounterOffer"), "{0}", getofertsXLngStr("DtxtCounter")), "{1}", txtOfert) & "</font>"
                  Case "R"
                  	Response.write "<font color=""#FF0066"">" & getofertsXLngStr("DtxtReject") & "</font>"
                  Case "C"
                  	Response.write "<font color=""#666699"">" & getofertsXLngStr("DtxtAnuled") & "</font>"
                  End Select %></td>
			</tr>
			<tr class="GeneralTbl">
				<td>
				<p align="center">
				<a href="javascript:doMyLink('ofertAgentContraOfert.asp', 'ofertIndex=<%=rs("ofertIndex")%>&status=<%=rs("ofertStatus")%><%=fltParam%>&page=<%=iPageCurrent%>&redir=<%=Request("cmd")%>', '');"><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></a></td>
				<td><% If 1 = 2 Then %>Oferta<% Else %><%=txtOfert%><% End If %>&nbsp;#<%=rs("OfertIndex")%></td>
				<td><% If rs("ofertDate") <> "" then %><%=FormatDate(rs("ofertDate"),True)%><% end if %>&nbsp;</td>
				<td><% If rs("ofertDueDate") <> "" then %><%=FormatDate(rs("ofertDueDate"),True)%><% end if %>&nbsp;</td>
				<td>
				<p align="right"><%=ofertQuantity%>&nbsp;</td>
				<td>
				<p align="right"><nobr><%=rs("Currency")%>&nbsp;<%=FormatNumber(ofertPrice,myApp.PriceDec)%></nobr></td>
				<td>
				<p align="right"><nobr>%&nbsp;<%=FormatNumber(rs("ofertCostPercentage"),myApp.PercentDec)%></nobr></td>
			</tr>
			<tr class="GeneralTbl">
				<td>
				&nbsp;</td>
				<td>
				<%=getofertsXLngStr("LtxtResp")%></td>
				<td><% If rs("responseDate") <> "" then %><%=FormatDate(rs("responseDate"),False)%><% end if %>&nbsp;</td>
				<td><% If rs("responseDueDate") <> "" then %><%=FormatDate(rs("responseDueDate"),False)%><% end if %>&nbsp;</td>
				<td>
				<p align="right"><%=responseQuantity%>&nbsp;</td>
				<td>
				<p align="right"><nobr><%=rs("Currency")%>&nbsp;<%=FormatNumber(responsePrice,myApp.PriceDec)%></nobr></td>
				<td>
				<p align="right"><nobr>%&nbsp;<%=FormatNumber(rs("responseCostPercentage"),myApp.PercentDec)%></nobr></td>
			</tr>
			<tr class="GeneralTbl">
				<td colspan="7"><hr size="1"></td>
			</tr>
		  <% rs.MoveNext
		  if rs.eof then exit for
		  next
		  else %>
			<tr class="GeneralTblBold2">
				<td colspan="7">
				<p align="center"><%=getofertsXLngStr("DtxtNoData")%></td>
			</tr>
			<% End If %>
		</table>
		</td>
	</tr>
	<% If iPageCount > 1 Then %>
	<tr>
		<td><% doOfertsXPages %></td>
	</tr>
	<% End If %>
</table>
<script language="javascript">
function GoLogView(CardCode) 
{
	document.viewLogNum.action = 'addCard/crdConfDetailOpen.asp';
	document.viewLogNum.target = '_blank';
	document.viewLogNum.CardCode.value = CardCode 
	document.viewLogNum.submit() 
}
function goViewItem(Item) 
{ 
	openItemDetails(Item);
}
</script>
<!--#include file="../itemDetails.inc"-->
<form target="_blank" method="post" name="viewLogNum" action="addCard/crdConfDetailOpen.asp">
<input type="hidden" name="CardCode" value="">
<input type="hidden" name="DocType" value="2">
<input type="hidden" name="ViewOnly" value="Y">
<input type="hidden" name="pop" value="Y">
<input type="hidden" name="AddPath" value="../">
<input type="hidden" name="Item" value="">
<input type="hidden" name="T1" value="">
<input type="hidden" name="cmd" value="">
</form>
<form name="frmGoP" action="<%=strScriptName%>">
<% For each itm in Request.Form
If itm <> "page" Then  %>
<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
<% 
End If 
Next %>
<input type="hidden" name="page" value="<%=iPageCurrent%>">
</form>
<script language="javascript">
function goP(p) { document.frmGoP.page.value = p; document.frmGoP.submit(); }
</script>
<% Sub doOfertsXPages %>
<table cellpadding="0" border="0" width="100%">
	<tr class="FirmTlt3">
		<% If iPageCurrent > 1 Then	%><td width="15">
		<p align="center">
		<a href='#' onclick="goP(<%= iPageCurrent - 1 %>);"><img border="0" src="design/0/images/<%=Session("rtl")%>prev_icon_trans.gif" width="15" height="15"></a></td><% End If %>
		<td>
		<p align="center">
		<% if iPageCount > 1 then
	    For I = 1 To iPageCount
		If I = iPageCurrent Then %><b><font size="3"><%= I %></font></b>
		<% Else %>
		<a class="LnkSearchPaginacion" href="#" onclick="goP(<%= I %>);"><%= I %></a>
		<% End If
		Next 'I
		end if %></td>
		<% If iPageCurrent < iPageCount Then %>
		<td width="15">
		<p align="center">
		<a href='#' onclick="goP(<%= iPageCurrent + 1 %>);"><img border="0" src="design/0/images/<%=Session("rtl")%>next_icon_trans.gif" width="15" height="15"></a></td><% End If %>
	</tr>
</table>
<% End Sub %>