<%@ Language=VBScript %>
<!--#include file="../chkLogin.asp" -->
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>

<!--#include file="lang/olkItemReport2.asp" -->
<!--#include file="../myHTMLEncode.asp" -->
<!--#include file="../authorizationClass.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=getolkItemReport2LngStr("LttlLastSal")%></title>
<script language="javascript">
function Start(page, w, h, s) {
OpenWin = this.open(page, "ImageThumb", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable=yes, width="+w+",height="+h);
}
</script>
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css">
</head>

<body SCROLL="no" topmargin="0" leftmargin="0">
<%

Dim myAut
set myAut = New clsAuthorization

set rs = Server.CreateObject("ADODB.recordset")
sql = 	"select top 1 OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', ItemCode, ItemName) ItemName, " & _
		"Replace(Convert(nvarchar(4000),IsNull(OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'UserText', ItemCode, UserText), '')),Char(13),'<br>') UserText, PicturName, " & _
		"(select olkItemReport2 from olkcommon) olkItemRep2 " & _
		"from oitm " & _
		"cross join oadm where itemcode = N'" & saveHTMLDecode(Request("Item"), False) & "' " & _
		"order by CurrPeriod desc"
set rs = conn.execute(sql)

If rs("olkItemRep2") = "D" Then repPrice = "Price" Else repPrice = "OLKCommon.dbo.DBOLKCode" & Session("ID") & "('" & rs("olkItemRep2") & "', Price, " & myApp.PriceDec & ") As 'Price'"
If rs("PicturName") <> "" Then Pic = rs("PicturName") Else Pic = "n_a.gif"

iPageSize = 13
set rw = server.CreateObject("ADODB.recordset")
If Request("page") = "" Then
	iPageCurrent = 1
Else
	iPageCurrent = CInt(Request("page"))
End If
If Request("order1") = "" Then order1 = "DocDateSort" Else order1 = Request("order1")
If Request("order2") = "" Then order2 = "desc" Else order2 = Request("order2")

sql = 	"select oinv.DocEntry, DocNum, LineNum, Convert(int,oinv.DocDate), OINV.DocDate, Convert(int,oinv.DocDate) DocDateSort, " & _
		"oinv.CardCode, IsNull(Left(CardName, 40), '') + Case When Len(CardName) > 40 Then '...' Else '' End CardName, Quantity, inv1.Currency,  " & repPrice & ", " & _
		"IsNull(Case UseBaseUn When 'N' Then OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'SalUnitMsr', inv1.ItemCode, SalUnitMsr) When 'Y' Then N'" & getolkItemReport2LngStr("DtxtUnit") & "' End, '') MType, OITM.NumInSale, UseBaseUn " & _
		"from inv1 " & _
		"inner join oinv on oinv.docentry = inv1.docentry " & _
		"inner join oitm on oitm.itemcode = inv1.itemcode " & _
		"where inv1.ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "' and (targettype is null or targettype <> 14) "
		
If Not myAut.HasAuthorization(97) Then sql = sql & " and oinv.SlpCode = " & Session("vendid") & " "
		
sql = sql & "order by " & order1 & " " & order2
	  	rw.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
		  	
rw.PageSize = iPageSize
rw.CacheSize = iPageSize
iPageCount = rw.PageCount
%>
<form method="post" action="olkItemReport2.asp" name="frmGoPage">
<% For each itm in Request.Form
If itm <> "page" and itm <> "iNextCount" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
<% End If
Next %>
<% For each itm in Request.QueryString
If itm <> "page" and itm <> "iNextCount" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
<% End If
Next %>
<% If Request("iNextCount") <> "" Then iNextCount = CInt(Request("iNextCount"))+1 Else iNextCount = 1 %>
<input type="hidden" name="iNextCount" value="<%=iNextCount%>">
<input type="hidden" name="page" value="">
<% If Request("iNextCount") = "" Then %>
<input type="hidden" name="order1" value="">
<input type="hidden" name="order2" value="">
<% End If %>
</form>
<script language="javascript">
<!--
function goPage(p)
{
	document.frmGoPage.page.value = p;
	document.frmGoPage.submit();
}
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
	document.frmGoPage.page.value = 1;
	document.frmGoPage.submit();
}
//-->
</script>
<!--#include file="../loadAlterNames.asp"-->
	<table border="0" cellpadding="0" width="594" id="table1">
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%" id="table2">
				<tr class="GeneralTblBold2">
					<td width="93"><%=getolkItemReport2LngStr("DtxtCode")%>:</td>
					<td><%=Request("Item")%>: <%=RS("ItemName")%></td>
				</tr>
				<tr class="GeneralTbl">
					<td width="93">
					<p align="center"><% If Pic <> "n_a.gif" then %><a href="javascript:Start('../thumb/?item=<%=saveHTMLDecode(Request("Item"), False)%>&pop=Y&AddPath=../',529,510,'yes')"><% end if %><img border="0" src="../pic.aspx?filename=<%=Pic%>&dbName=<%=Session("olkdb")%>"><% If Pic <> "n_a.gif" then %></a><% end if %></td>
					<td valign="top">
					<ilayer name="scroll1" width=100% height=100 clip="0,0,170,150">
					<layer name="scroll2" width=100% height=100 bgColor="white">
					<div id="scroll3" style="width:100%;height:100px;background-color:white;overflow:auto">
					<%=vartext%><% If Not IsNull(rs("UserText")) Then %><%=rs("UserText")%><% End If %>
					</div>
					</layer>
					</ilayer>
					</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr class="GeneralTbl">
			<td><hr size="1"></td>
		</tr>
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%" id="table4">
			<% If iPageCount > 1 Then %>
				<tr class="GeneralTbl">
					<td colspan="7">
					<table border="0" cellpadding="0" width="100%" id="table5" cellspacing="0">
						<tr>
							<td width="12"><% If iPageCurrent > 1 Then	%><a href="javascript:goPage(<%= iPageCurrent - 1 %>);"><img border="0" src="../design/0/images/<%=Session("rtl")%>prev_icon.gif"></a><% End If %></td>
							<td align="center"><div class="LinkNormal"><%
							For i = 1 to iPageCount %>
							<% If iPageCurrent <> i Then %><a class="LinkTop" href="javascript:goPage(<%=i%>);"><% Else %><font size="3"><% End If %>
							<%=i%>
							<% If iPageCurrent <> i Then %></a><% Else %></font><% End If %>&nbsp;
							<% Next %></div></td>
							<td align="right" width="12"><% If iPageCurrent < iPageCount Then %><a href="javascript:goPage(<%= iPageCurrent + 1 %>);"><img border="0" src="../design/0/images/<%=Session("rtl")%>next_icon.gif"></a><% End If %></td>
						</tr>
					</table>
					</td>
				</tr>
				<% End If %>
				<tr class="GeneralTblBold2">
					<td align="center" colspan="2" style="cursor: hand" onclick="javascript:doSort('DocNum');" <% doItemRepSortBG("DocNum")%>>#<% doItemRepSortImg("DocNum")%></td>
					<td align="center" style="cursor: hand" onclick="javascript:doSort('DocDateSort');" <% doItemRepSortBG("DocDateSort")%>><%=getolkItemReport2LngStr("DtxtDate")%><% doItemRepSortImg("DocDateSort")%></td>
					<td align="center" style="cursor: hand" onclick="javascript:doSort('CardName');" <% doItemRepSortBG("CardName")%>><%=txtClient%><% doItemRepSortImg("CardName")%></td>
					<td align="center" style="cursor: hand" onclick="javascript:doSort('Quantity');" <% doItemRepSortBG("Quantity")%>><%=getolkItemReport2LngStr("LtxtQty")%><% doItemRepSortImg("Quantity")%></td>
					<td align="center" style="cursor: hand" onclick="javascript:doSort('MType');" <% doItemRepSortBG("MType")%>><%=getolkItemReport2LngStr("LtxtSalMet")%><% doItemRepSortImg("MType")%></td>
					<td align="center" style="cursor: hand" onclick="javascript:doSort('Price');" <% doItemRepSortBG("Price")%>><%=getolkItemReport2LngStr("DtxtPrice")%><% doItemRepSortImg("Price")%></td>
				</tr>
				<% If iPageCount > 0 Then
				rw.AbsolutePage = iPageCurrent
				For i = 1 to rw.PageSize %>
				<tr class="GeneralTbl">
					<td>
					<a href="javascript:goDetail(<%=rw("LineNum")%>,<%=rw("DocEntry")%>)"><img border="0" src="../design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></a></td>
					<td><%=RW("DocNum")%>(<%=RW("LineNum")+1%>)</td>
					<td><%=FormatDate(RW("DocDate"), True)%>&nbsp;</td>
					<td><%=RW("CardName")%>&nbsp;</td>
					<td align="right"><%=FormatNumber(RW("Quantity"),myApp.QtyDec)%>&nbsp;</td>
					<td><%=RW("MType")%><% If rw("UseBaseUn") = "N" and myApp.GetShowQtyInUn Then %>(<%=rw("NumInSale")%>)<% End If %>&nbsp;</td>
					<td align="right"><% If rs("olkItemRep2") = "D" Then %><nobr><%=Rw("Currency")%>&nbsp;<%=FormatNumber(RW("Price"),myApp.PriceDec)%></nobr><% Else %><%=rw("Price")%><% End If %></td>
				</tr>
				<% rw.movenext
				If rw.eof then exit for
				Next
				Else %>
				<tr class="GeneralTbl">
					<td colspan="7">
					<p align="center"><%=getolkItemReport2LngStr("DtxtNoData")%></td>
				</tr>
				<% End If %>
				<tr class="GeneralTbl">
					<td colspan="2">&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr class="GeneralTbl">
			<td>
			<input type="button" value="<%=getolkItemReport2LngStr("DtxtBack")%>" name="B1" style="float: <% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>" onclick="javascript:history.go(-<%=iNextCount%>);"></td>
		</tr>
	</table>
</div>
<form name="frmDetail" method="post" target="_blank" action="../cxcDocDetailOpen.asp">
<input type="hidden" name="high" value="">
<input type="hidden" name="doctype" value="13">
<input type="hidden" name="DocEntry" value="">
<input type="hidden" name="order1" value="">
<input type="hidden" name="order2" value="">
</form>
<script language="javascript">
<!--
function goDetail(high, DocEntry) {
	document.frmDetail.high.value = high;
	document.frmDetail.DocEntry.value = DocEntry;
	document.frmDetail.submit();
}
//-->
</script>
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