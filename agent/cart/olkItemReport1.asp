<%@ Language=VBScript %>
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->

<!--#include file="lang/olkItemReport1.asp" -->
<!--#include file="../myHTMLEncode.asp" -->
<!--#include file="../authorizationClass.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<% If Request("view") = "OLK" Then ttlBy = "OLK" Else ttlBy = getolkItemReport1LngStr("LtxtSAPOrders") %>
<title><%=Replace(getolkItemReport1LngStr("LttlResItms"), "{0}", ttlBy)%></title>
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css">
<style>
.nobr
{
	white-space: nowrap;
}
</style>
</head>
<script language="javascript">
function Start(page, w, h, s) {
OpenWin = this.open(page, "ImageThumb", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable=yes, width="+w+",height="+h);
}
</script>
<body SCROLL="no" topmargin="0" leftmargin="0" link="#6EB7FC" vlink="#6EB7FC">
<%

Dim myAut
set myAut = New clsAuthorization


iPageSize = 14 %>
<!--#include file="../loadAlterNames.asp" -->
<%
      set rs = Server.CreateObject("ADODB.recordset")
sql = "select top 1 OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', ItemCode, ItemName) ItemName, " & _
	"Replace(Convert(nvarchar(4000),IsNull(OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'UserText', ItemCode, UserText), '')),Char(13),'<br>') UserText, " & _
	"PicturName " & _
	"from oitm cross join oadm where itemcode = N'" & saveHTMLDecode(Request("Item"), False) & "' order by CurrPeriod desc"
set rs = conn.execute(sql)
If myApp.olkItemReport2 = "D" Then repPrice = "Price" Else repPrice = "OLKCommon.dbo.DBOLKCode" & Session("ID") & "('" & myApp.olkItemReport2 & "', Price, " & myApp.PriceDec & ") As 'Price'"


		  If rs("PicturName") <> "" Then
			  Pic = rs("PicturName")
		  Else
			  Pic = "n_a.gif"
		  End If 
		  
		  If Request("order1") = "" Then order1 = "DocDateNum" Else order1 = Request("order1")
		  If Request("order2") = "" Then order2 = "desc" Else order2 = Request("order2")
		  
		  If Not myAut.HasAuthorization(97) Then
		  	slpOLKFilter = " and tdoc.SlpCode = " & Session("vendid") & " "
		  	slpSBOFilter = " and T0.SlpCode = " & Session("vendid") & " "
		  End If
		  
		  set rw = server.CreateObject("ADODB.recordset")
		  If Request("view") = "OLK" Then
			  sql = "select -2 ObjectCode, tlog.Object ObjectCodeType, tlog.LogNum, tlog.LogNum DocNum, doc1.LineNum, DocDate, " & _
			  		"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', tdoc.SlpCode, SlpName) SlpName, " & _
		  			"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'SalUnitMsr', doc1.ItemCode, SalUnitMsr) SalUnitMsr, " & _
		  			"IsNull(Case SaleType When 1 Then N'" & getolkItemReport1LngStr("DtxtUnit") & "' When 2 Then OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'SalUnitMsr', doc1.ItemCode, SalUnitMsr) When 3 Then OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'SalPackMsr', doc1.ItemCode, SalPackMsr) End, '') As SaleType, " & _
		  			"IsNull(Case SaleType When 2 Then NumInSale When 3 Then SalPackUn End, 1) SaleTypeNum, " & _
		  			"IsNull(tdoc.CardName, tdoc.CardCode) CardCode, " & _
			  		"IsNull(CardName, '') CardName, Quantity, " & repPrice & " , doc1.Currency, SalPackUn, SaleType SaleType2, NumInSale, Convert(int,DocDate) DocDateNum " & _
			  		"from r3_obscommon..doc1 doc1 " & _
			  		"inner join r3_obscommon..tdoc tdoc on tdoc.lognum = doc1.lognum " & _
			  		"inner join r3_obscommon..tlog tlog on tlog.lognum = tdoc.lognum " & _
			  		"left outer join OLKSalesLines T0 on T0.LogNum = doc1.Lognum and T0.LineNum = doc1.Linenum " & _
			  		"inner join oslp on oslp.slpcode = tdoc.slpcode " & _
			  		"inner join oitm on oitm.itemcode = doc1.itemcode collate database_default " & _
			  		"inner join R3_ObsCommon..TLOGControl X0 on X0.LogNum = doc1.LogNum " & _
			  		"where tlog.Object in (13,15,17) and doc1.ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "' " & slpOLKFilter & " and Status in ('R','H') and Company = db_name() and X0.appId = 'TM-OLK' " & _
			  		" order by " & order1 & " " & order2
		  ElseIf Request("view") = "ORDR" Then
			  sql = "select 17 ObjectCode, 17 ObjectCodeType, T0.DocEntry LogNum, T0.DocNum, T1.LineNum, T0.DocDate, " & _
		  			"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', T0.SlpCode, SlpName) SlpName, " & _
		  			"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'SalUnitMsr', T1.ItemCode, SalUnitMsr) SalUnitMsr, " & _
		  			"IsNull(Case T1.UseBaseUn When 'Y' Then N'" & getolkItemReport1LngStr("DtxtUnit") & "' When 'N' Then OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'SalUnitMsr', T1.ItemCode, SalUnitMsr) End, '') As SaleType, " & _
		  			"NumInSale SaleTypeNum, " & _
		  			"Left(IsNull(T0.CardName,T0.CardCode), 22) + Case When Len(T0.CardName) > 22 Then '...' Else '' End CardCode, IsNull(T0.CardName, '') CardName, " & _
			  		"T1.OpenQty Quantity, " & repPrice & " , T1.Currency, " & _
			  		"SalPackUn, Case T1.UseBaseUn When 'Y' Then 1 Else 2 End SaleType2, NumInSale, Convert(int, T0.DocDate) DocDateNum " & _
			  		"from ORDR T0 " & _
			  		"inner join RDR1 T1 on T1.DocEntry = T0.DocEntry " & _
			  		"inner join OSLP T2 on T2.SlpCode = T0.SlpCode " & _
			  		"inner join OITM T3 on T3.ItemCode = T1.ItemCode " & _
			  		"where T1.ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "' and T0.DocStatus = 'O' and T1.LineStatus = 'O' " & slpSBOFilter & "  " & _
			  		"union all " & _
			  		"select 13 ObjectCode, 13 ObjectCodeType, T0.DocEntry LogNum, T0.DocNum, T1.LineNum, T0.DocDate, " & _
			  		"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', T0.SlpCode, SlpName) SlpName, " & _
			  		"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'SalUnitMsr', T1.ItemCode, SalUnitMsr) SalUnitMsr, " & _
		  			"IsNull(Case T1.UseBaseUn When 'Y' Then N'" & getolkItemReport1LngStr("DtxtUnit") & "' When 'N' Then OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'SalUnitMsr', T1.ItemCode, SalUnitMsr) End, '') As SaleType, " & _
		  			"NumInSale SaleTypeNum, " & _
		  			"Left(IsNull(T0.CardName,T0.CardCode), 22) + Case When Len(T0.CardName) > 22 Then '...' Else '' End CardCode, IsNull(T0.CardName, '') CardName, " & _
					"T1.OpenQty Quantity, " & repPrice & " , T1.Currency, " & _
					"SalPackUn, Case T1.UseBaseUn When 'Y' Then 1 Else 2 End SaleType2, NumInSale, Convert(int, T0.DocDate) DocDateNum " & _
					"from OINV T0 " & _
					"inner join INV1 T1 on T1.DocEntry = T0.DocEntry " & _
					"inner join OSLP T2 on T2.SlpCode = T0.SlpCode " & _
					"inner join OITM T3 on T3.ItemCode = T1.ItemCode " & _
					"where T1.ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "' and T0.DocStatus = 'O' and T1.LineStatus = 'O' and T0.IsIns = 'Y' " & slpSBOFilter & "  " & _
			  		"order by " & order1 & " " & order2
		  End If
		  rw.pagesize = iPageSize
		  rw.cachesize = iPageSize
		  rw.open sql, conn, 3, 1
		  If Request("page") = "" Then 
		  	iCurrent = 1
		  	varBack = 1
		  Else 
		  	iCurrent = Request("page")
		  	varBack = Request("Back") + 1
		  End If
		  If rw.recordcount > 0 Then rw.AbsolutePage = iCurrent
%>
<form method="post" action="olkItemReport1.asp" name="frmGoPage">
<% For each itm in Request.Form
If itm <> "Page" and itm <> "Back" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
<% End If
Next %>
<% For each itm in Request.QueryString
If itm <> "Page" and itm <> "Back" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
<% End If
Next %>
<% If Request("Back") <> "" Then Back = CInt(Request("Back"))+1 Else Back = 1 %>
<input type="hidden" name="Back" value="<%=Back%>">
<input type="hidden" name="Page" value="">
<% If Request("Back") = "" Then %>
<input type="hidden" name="order1" value="">
<input type="hidden" name="order2" value="">
<% End If %>
</form>
<script language="javascript">
<!--
function doSort(c)
{
	document.frmGoPage.order1.value = c;
	if ('<%=order1%>' == c)
	{
		if ('<%=order2%>' == 'asc')
			document.frmGoPage.order2.value = 'desc';
		else
			document.frmGoPage.order2.value = 'asc';
	}
	else
	{
		document.frmGoPage.order2.value = 'asc';
	}
	document.frmGoPage.Page.value = 1;
	document.frmGoPage.submit();
}
function goPage(p)
{
	document.frmGoPage.Page.value = p;
	document.frmGoPage.submit();
}
//-->
</script>
<div align="left">
	<table border="0" cellpadding="0" width="594" id="table1">
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%" id="table2">
				<tr class="GeneralTblBold2">
					<td width="93"><%=getolkItemReport1LngStr("DtxtCode")%>:</td>
					<td><%=Request("Item")%>: <%=RS("ItemName")%></td>
				</tr>
				<tr class="GeneralTbl">
					<td width="93">
					<p align="center"><a href="javascript:Start('../thumb/?item=<%=saveHTMLDecode(Request("Item"), False)%>&pop=Y&AddPath=../',529,510,'yes')"><img border="0" src="../pic.aspx?filename=<%=Pic%>&dbName=<%=Session("olkdb")%>"></a></td>
					<td valign="top">
					<div id="scroll3" style="width: 491px;height:100px;background-color:white;overflow:auto">
					<%=vartext%><% If Not IsNull(rs("UserText")) Then %><%=rs("UserText")%><% End If %>
					</div>
					</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr class="GeneralTbl">
			<td><hr size="1"></td>
		</tr>
		<% If rw.PageCount > 1 Then %>
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%" cellspacing="1" id="table5">
				<tr class="GeneralTbl">
					<td width="19">
					<p align="center"><div class="LinkNormal">
					<% If CInt(iCurrent) > 1 Then %><a href="javascript:goPage(<%=iCurrent-1%>);"><img border="0" src="../design/0/images/<%=Session("rtl")%>prev_icon_trans.gif" width="15" height="15"></a><% End If %></td>
					<td><p align="center">
					<% For i = 1 to rw.PageCount %>
					<% If CInt(i) <> CInt(iCurrent) Then %><a class="LinkTop" href="javascript:goPage(<%=i%>);"><% Else %><font size="3"><% End If %>
					<%=i%>
					<% If CInt(i) <> CInt(iCurrent) Then %></a><% Else %></font><% End If %>&nbsp;
					<% Next %></p></td>
					<td width="16">
					<p align="center">
					<% If CInt(iCurrent) < rw.PageCount Then %><a href="javascript:goPage(<%=iCurrent+1%>);"><img border="0" src="../design/0/images/<%=Session("rtl")%>next_icon_trans.gif" width="15" height="15"></a><% End If %></td>
				</tr>
			</table>
			</td>
		</tr>
		<% End If %>
		<tr>
			<td>
		<div id="scrollDetail3" style="width:590px;height:240px;background-color:white;overflow:auto">
			<table border="0" cellpadding="0" width="100%">
				<tr class="GeneralTblBold2">
					<td align="center" colspan="2" style="cursor: hand" onclick="javascript:doSort('DocNum');" <% doItemRepSortBG("DocNum")%>>#<% doItemRepSortImg("DocNum")%></td>
					<td align="center" style="cursor: hand" onclick="javascript:doSort('ObjectCodeType');" <% doItemRepSortBG("ObjectCodeType")%>><%=getolkItemReport1LngStr("DtxtType")%><% doItemRepSortImg("ObjectCodeType")%></td>
					<td align="center" style="cursor: hand" onclick="javascript:doSort('DocDateNum');" <% doItemRepSortBG("DocDateNum")%>><%=getolkItemReport1LngStr("DtxtDate")%><% doItemRepSortImg("DocDateNum")%></td>
					<td align="center" style="cursor: hand" onclick="javascript:doSort('SlpName');" <% doItemRepSortBG("SlpName")%>><% If 1 = 2 Then %>Vendedor<% Else %><%=txtAgent%><% End If %><% doItemRepSortImg("SlpName")%></td>
					<td align="center" style="cursor: hand" onclick="javascript:doSort('CardName');" <% doItemRepSortBG("CardName")%>><% If 1 = 2 Then %>Cliente<% Else %><%=txtClient%><% End If %><% doItemRepSortImg("CardName")%></td>
					<td align="center" style="cursor: hand" onclick="javascript:doSort('Quantity');" <% doItemRepSortBG("Quantity")%>><%=getolkItemReport1LngStr("LtxtQty")%><% doItemRepSortImg("Quantity")%></td>
					<td align="center" style="cursor: hand" onclick="javascript:doSort('SaleType');" <% doItemRepSortBG("SaleType")%>><%=getolkItemReport1LngStr("LtxtUn")%><% doItemRepSortImg("SaleType")%></td>
					<td align="center" style="cursor: hand" onclick="javascript:doSort('Price');" <% doItemRepSortBG("Price")%>><%=getolkItemReport1LngStr("DtxtPrice")%><% doItemRepSortImg("Price")%></td>
				</tr>
				<% If rw.recordcount > 0 Then
				For i = 1 to rw.pagesize %>
				<tr class="GeneralTbl" style="">
					<td width="20"><a href="javascript:goViewDoc(<%=Rw("ObjectCode")%>, <%=Rw("LogNum")%>, <%=rw("LineNum")%>);"><img border="0" src="../design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></a></td>
					<td class="nobr"><%=RW("DocNum")%><% If Request("view") = "ORDR" Then %>(<%=rw("LineNum")+1%>)<% End If %>&nbsp;</td>
					<td class="nobr"><%
						Select Case RW("ObjectCodeType")
							Case 13
								Response.Write txtInv
							Case 17
								Response.Write txtOrdr
						End Select %></td>
					<td class="nobr"><%=FormatDate(RW("DocDate"), True)%>&nbsp;</td>
					<td class="nobr"><%=RW("SLPName")%>&nbsp;</td>
					<td class="nobr"><%=RW("CardCode")%>&nbsp;</td>
					<td align="right" class="nobr">&nbsp;<% If rw("SaleType2") = 3 Then %><%=FormatNumber(CDbl(RW("Quantity"))/CDbl(RW("SalPackUn")),myApp.QtyDec)%><% Else %><%=FormatNumber(RW("Quantity"),myApp.QtyDec)%><% End If %></td>
					<td align="center" class="nobr">
					<%=RW("SaleType")%><% If (rw("SaleType2") = 2 or rw("SaleType2") = 3) and myApp.GetShowQtyInUn Then %>(<%=rw("SaleTypeNum")%>)<% End If %><% If Not myApp.UnEmbPriceSet And rw("SaleType2") = 3 Then %>&nbsp;<%=rw("SalUnitMsr")%><% If myApp.GetShowQtyInUn Then %>(<%=rw("NumInSale")%>)<% End If %><% End If %></td>
					<td align="right" class="nobr">&nbsp;<% If myApp.olkItemReport2 = "D" Then %><nobr><%=RW("Currency")%>&nbsp;<% If myApp.UnEmbPriceSet And rw("SaleType2") = 3 Then %><%=FormatNumber(CDbl(RW("Price"))*CDbl(rw("SalPackUn")),myApp.PriceDec)%><% Else %><%=FormatNumber(RW("Price"),myApp.PriceDec)%><% End If %></nobr><% Else %><%=rw("Price")%><% End If %></td>
				</tr>
		      <% rw.movenext
		      If rw.eof then exit for
		      next
		      Else %>
				<tr class="GeneralTbl">
					<td colspan="9" align="center"><%=getolkItemReport1LngStr("DtxtNoData")%></td>
				</tr>
		      <% End If %>
			</table>
			</div>
			</td>
		</tr>
		<tr class="GeneralTbl">
			<td>
			<input type="submit" value="<%=getolkItemReport1LngStr("DtxtBack")%>" name="B1" style="float: <% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>" onclick="javascript:history.go(-<%=varBack%>);"></td>
		</tr>
	</table>
</div>
<script language="javascript">
function goViewDoc(ObjCode, DocEntry, LineNum)
{
	document.frmViewDet.DocType.value = ObjCode;
	document.frmViewDet.DocEntry.value = DocEntry;
	document.frmViewDet.high.value = LineNum;
	document.frmViewDet.submit();
}
</script>
<form target="_blank" method="post" name="frmViewDet" action="../cxcDocDetailOpen.asp">
<input type="hidden" name="DocEntry" value="">
<input type="hidden" name="high" value="">
<input type="hidden" name="DocType" value="">
<input type="hidden" name="pop" value="Y">
<input type="hidden" name="AddPath" value="">
</form>
<% 
Sub doItemRepSortImg(c)
	If LCase(order1) = LCase(c) Then
		If order2 = "asc" Then
			Response.Write "<img src=""../images/arrow_up.gif"">"
		Else
			Response.Write "<img src=""../images/arrow_down.gif"">"
		End If
	End If
End Sub 
Sub doItemRepSortBG(c)
	If LCase(order1) = LCase(c) Then Response.Write "class=""GeneralTblBold2HighLight"""
End Sub %>
  </body>

</html>