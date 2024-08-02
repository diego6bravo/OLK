<% addLngPathStr = "ventas/" %>
<!--#include file="lang/ventasX.asp" -->
<%
set rs = Server.CreateObject("ADODB.recordset")
AsignedSlp = not myAut.HasAuthorization(60)

ObjCode = ""

If CardType <> "S" or strScriptName <> "activeclient.asp" Then
	If myApp.EnableOQUT Then ObjCode = "23"
	If myApp.EnableORDR Then ObjCode = myAut.ConcValue(ObjCode, "17")
	If myApp.EnableODLN Then ObjCode = myAut.ConcValue(ObjCode, "15")
	If myApp.EnableODPIReq Then ObjCode = myAut.ConcValue(ObjCode, "203")
	If myApp.EnableODPIInv Then ObjCode = myAut.ConcValue(ObjCode, "204")
	If myApp.EnableOINV or myApp.EnableCashInv or myApp.EnableOINVRes Then ObjCode = myAut.ConcValue(ObjCode, "13")
	If myApp.EnableORCT Then ObjCode = myAut.ConcValue(ObjCode, "24")
End If

If CardType = "S" or strScriptName <> "activeclient.asp" Then
	If myApp.EnableOPOR Then ObjCode = myAut.ConcValue(ObjCode, "22")
End If

If ObjCode <> "" or myApp.EnableORCT Then

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKSeachOpenDocs" & Session("ID")
cmd.Parameters.Refresh()
cmd("@AllowAgentAccessCDoc") = GetYN(myApp.AllowAgentAccessCDoc)
cmd("@LanID") = Session("LanID")
cmd("@SlpCode") = Session("vendid")

If strScriptName = "activeclient.asp" Then 
	cmd("@CardCodeFrom") = Session("UserName")
	cmd("@CardCodeTo") = Session("UserName")
Else
	If Request("CardCodeFrom") <> "" Then cmd("@CardCodeFrom") = Request("CardCodeFrom")
	If Request("CardCodeTo") <> "" Then cmd("@CardCodeTo") = Request("CardCodeTo")
End If

If Request("ItemCodeFrom") <> "" Then cmd("@ItemCodeFrom") = Request("ItemCodeFrom")
If Request("ItemCodeTo") <> "" Then cmd("@ItemCodeTo") = Request("ItemCodeTo")

If Request("LogNumFrom") <> "" Then cmd("@LogNumFrom") = Request("LogNumFrom")
If Request("LogNumTo") <> "" Then cmd("@LogNumTo") = Request("LogNumTo") 

cmd("@All") = "Y"

If AsignedSlp or not myAut.HasAuthorization(97) Then 
	cmd("@All") = "N"
End If

If Request("Comments") <> "" Then cmd("@Comments") = Request("Comments")

If Request("GroupNameFrom") <> "" Then cmd("@GroupNameFrom") = Request("GroupNameFrom")
If Request("GroupNameTo") <> "" Then cmd("@GroupNameTo") = Request("GroupNameTo")

If Request("CountryFrom") <> "" Then cmd("@CountryNameFrom") = Request("CountryFrom")
If Request("CountryTo") <> "" Then cmd("@CountryNameTo") = Request("CountryTo")

If Request("CardString") <> "" Then cmd("@CardString") = Request("CardString")

If Request("DocType") <> "" Then cmd("@DocType") = Request("DocType")

If Request("SlpCodeFrom") <> "" Then cmd("@SlpNameFrom") = Request("SlpCodeFrom")
If Request("SlpCodeTo") <> "" Then cmd("@SlpNameTo") = Request("SlpCodeTo")

If Request("dtFrom") <> "" Then cmd("@DateFrom") = SaveCmdDate(Request("dtFrom"))
If Request("dtTo") <> "" Then cmd("@DateTo") = SaveCmdDate(Request("dtTo"))

If Request("orden1") <> "" Then 
	cmd("@Order") = Request("orden1")
	orden1 = Request("orden1")
Else
	orden1 = "0"
End If
If Request("orden2") <> "" Then 
	cmd("@OrderDir") = Request("orden2")
	orden2 = Request("orden2")
Else
	orden2 = "D"
End If
cmd("@Objects") = ObjCode
cmd("@EnableInv") = GetYN(myApp.EnableOINV)
cmd("@EnableCashInv") = GetYN(myApp.EnableCashInv)
cmd("@EnableInvRes") = GetYN(myApp.EnableOINVRes)

set rx = Server.CreateObject("ADODB.RecordSet")
rx.CursorLocation = 3 ' adUseClient
rx.open cmd
rx.PageSize = 40
nPageCount = rx.PageCount
If Request("Page") <> "" Then nPage = CLng(Request("Page")) Else nPage = 1
'If nPage < 1 Or nPage > nPageCount Then	nPage = 1

iNextCount = nPage
iCurMax = nPageCount/15
iCurNext = 0
do while iNextCount > 0
iNextCount = iNextCount - 15
iCurNext = iCurNext + 1
loop
If iCurMax - CInt(iCurMax) > 0 Then iCurMax = CInt(iCurMax) + 1

fromI = (iCurNext*15)-14
toI = (iCurNext*15)

If iCurMax <= iCurNext Then toI = nPageCount
If nPage > nPageCount Then nPage = nPageCount
If nPage < 1 Then nPage = 1
If Not rx.Eof then rx.AbsolutePage = nPage

%>
<script language="javascript">
function listPendAlert(obj) {
var objType;
switch (obj) {
	case 15:
		objType = "<%=txtOdln%>";
		break;
	case 17:
		objType = "<%=txtOrdr%>";
		break;
	case 23:
		objType = "<%=txtQuote%>";
		break;
	case 24:
		objType = "<%=txtRct%>";
		break;
	case 48:
		objType = "<%=txtInv%>/<%=txtRct%>";
		break;
	case 13:
		objType = "<%=txtInv%>";
		break;
	case 4:
		objType = "<%=getventasXLngStr("DtxtItem")%>";
		break;
	case 2:
		objType= "<%=txtClient%>"
		break;
}
alert('<%=getventasXLngStr("LtxtDisObj")%>'.replace('{0}', objType));
}

function confReopen()
{
	return confirm('<%=getventasXLngStr("LtxtConfReOpen")%>')
}
function valFrm()
{
	if (document.frmVentasX.chkDel.length)
	{
		var found = false;
		for (var i = 0;i<document.frmVentasX.chkDel.length;i++)
		{
			if (document.frmVentasX.chkDel[i].checked)
			{
				found = true;
				break;
			}
		}
		if (!found)
		{
			alert('<%=getventasXLngStr("LtxtValSelDoc")%>');
			return false;
		}
	}
	else
	{
		if (!document.frmVentasX.chkDel.checked)
		{
			alert('<%=getventasXLngStr("LtxtValSelDoc")%>');
			return false;
		}
	}
	return confirm('<%=getventasXLngStr("LtxtCondDelDoc")%>');
}
</script>
<div align="center">
<table border="0" cellpadding="0" width="100%">
<form name="frmVentasX" method="post" action="ventas/docdel.asp" onsubmit="javascript:return valFrm();">
	<tr class="GeneralTlt">
		<td><%=getventasXLngStr("LttlPendDocs")%></td>
	</tr>
	<% If rx.PageCount > 1 Then %>
	<tr>
		<td align="center"><% doVentasXPages %></td>
	</tr>
	<% End If %>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr class="FirmTlt3">
				<td width="15" align="center">&nbsp;</td>
				<td align="center" style="width: 18px">&nbsp;</td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('0');" <% doVentasXSortBG("0")%>><%=getventasXLngStr("DtxtLogNum")%><% doVentasXSortImg("0")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('4');" <% doVentasXSortBG("4")%>><%=txtAgent%><% doVentasXSortImg("4")%></td>
				<% If strScriptName <> "activeclient.asp" Then %>
				<td align="center" style="width: 18px">&nbsp;</td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('1');" <% doVentasXSortBG("1")%>><%=getventasXLngStr("DtxtCode")%><% doVentasXSortImg("1")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('2');" <% doVentasXSortBG("2")%>><%=getventasXLngStr("DtxtName")%><% doVentasXSortImg("2")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('5');" <% doVentasXSortBG("5")%>><%=getventasXLngStr("DtxtGroup")%><% doVentasXSortImg("5")%></td>
				<td align="center" style="cursor: hand; width: 30px;" onclick="javascript:doSort('6');" <% doVentasXSortBG("6")%>><% doVentasXSortImg("6")%></td><% End If %>
				<td align="center" style="cursor: hand; width: 75px;" onclick="javascript:doSort('3');" <% doVentasXSortBG("3")%>><%=getventasXLngStr("DtxtDate")%><% doVentasXSortImg("3")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('7');" <% doVentasXSortBG("7")%>><%=getventasXLngStr("DtxtType")%><% doVentasXSortImg("7")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('8');" <% doVentasXSortBG("8")%>><%=getventasXLngStr("DtxtState")%><% doVentasXSortImg("8")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('9');" <% doVentasXSortBG("9")%>><%=getventasXLngStr("DtxtTotal")%><% doVentasXSortImg("9")%></td>
			</tr>
		  <%  if not rx.eof then
		  LogNum = ""
		  do while not (rx.eof Or rx.AbsolutePage <> nPage )
		  	If LogNum <> "" Then LogNum = LogNum & ", "
		  	LogNum = LogNum & rx("LogNum")
		  rx.movenext
		  loop 
		  
		  set cmd = Server.CreateObject("ADODB.Command")
		  cmd.ActiveConnection = connCommon
		  cmd.CommandType = &H0004
		  cmd.CommandText = "DBOLKSearchOpenDocsData" & Session("ID")
		  cmd.Parameters.Refresh()
		  cmd("@LanID") = Session("LanID")
		  cmd("@LogNum") = LogNum
		  
		  If Request("orden1") <> "" Then cmd("@Order") = Request("orden1")
		  If Request("orden2") <> "" Then cmd("@OrderDir") = Request("orden2")


		rs.open cmd, , 3, 1
		do while not rs.eof
		  Enable = True
		  Select Case rs("Object")
		  	Case 13
		  		If Not IsNULL(rs("PayLogNum")) Then
		  			Enable = myApp.EnableCashInv
		  		ElseIf rs("ReserveInvoice") = "Y" Then
		  			Enable = myApp.EnableOINVRes
		  		Else
		  			Enable = myApp.EnableOINV
		  		End If
		  	Case 15
		  			Enable = myApp.EnableODLN
		  	Case 17
		 			Enable = myApp.EnableORDR
		  	Case 23
		  			Enable = myApp.EnableOQUT
		  	Case 24
		  			Enable = myApp.EnableORCT 
		  	Case 203
		  			Enable = myApp.EnableODPIReq
		  	Case 204
		  			Enable = myApp.EnableODPIInv
		  End Select %>
			<tr class="<% If rs("Source") = "V" Then %>GeneralTbl<% Else %>CanastaTblExpense<% End If %>">
				<td width="15" align="center" style="height: 15px">
				<img src="images/checkbox_off.jpg" border="0" onclick="doCheckDel(this, <%=rs("LogNum")%>);">
				<input type="checkbox" name="chkDel" id="chkDel<%=rs("LogNum")%>" value="<%=rs("LogNum")%>" style="display: none;">
				</td>
				<td align="center" style="width: 18px; height: 15px;">
				<a href="javascript:<% If Enable Then %>doGoDoc(<%=rs("Object")%>, '<%=rs("LogNum")%>', '<%=rs("PayLogNum")%>', '<%=Replace(myHTMLEncode(rs("CardCode")), "'", "\'")%>', '<%=rs("status")%>');<% Else %>listPendAlert(<%=rs("Object")%>);<% End If %>">
				<img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></a></td>
				<td style="height: 15px"><%=RS("LogNum")%>&nbsp;</td>
				<td style="height: 15px"><%=RS("SlpName")%>&nbsp;</td>
				<% If strScriptName <> "activeclient.asp" Then %>
				<td align="center" style="width: 18px; height: 15px;">
				<a href="javascript:goDetail(2, '<%=Replace(myHTMLEncode(RS("CardCode")), "'", "\'")%>')">
				<img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" alt="<%=RS("CardCode")%>" width="15" height="13"></a></td>
				<td style="height: 15px"><% If Not isNull(rs("CardCode")) Then %><%=RS("CardCode")%><% End If %>&nbsp;</td>
				<td style="height: 15px"><% If Not isNull(rs("CardName")) Then %><%=RS("cardname")%><% End If %>&nbsp;</td>
				<td style="height: 15px"><% If Not isNull(rs("GroupName")) Then %><%=RS("GroupName")%><% End If %>&nbsp;</td>
				<td align="center" style="width: 30px; height: 15px;">
				<img src="images/country/pic.aspx?filename=<%=rs("Country")%>.gif&MaxHeight=15" alt="<%=rs("CountryName")%>">
				</td><% End If %>
				<td style="width: 75px; text-align: center; height: 15px;"><%=FormatDate(RS("DocDate"), True)%>&nbsp;</td>
				<td style="height: 15px"><%
			     Select Case RS("Object")
			     Case 17
			    	 Response.write txtOrdr
			     Case 23
			     	Response.write txtQuote
			     Case 24
			    	 Response.write txtRct
			     Case 13
			     	Select Case rs("ReserveInvoice")
			     		Case "Y"
					     	Response.write txtInvRes
			     		Case Else
					     	Response.write txtInv
					End Select
			     Case 14
			     	Response.write "Nota de credito"
			     Case 16
			     	Response.write "Devoluciones"
			     Case 18
			    	 Response.write "Comprobante de compra"
			     Case 20
			     	Response.write "Consignaci�n de mercancia"
			     Case 59
			     	Response.write "Entreda general al inventario"
			     Case 15
			     	Response.write txtOdln '"Entregas"
			     Case 19
			     	Response.write "Nota de debito"
			     Case 21
			     	Response.write "Devoluciones en compra"
			     Case 60
			    	 Response.write "Salida general del inventario"
			     Case 67
			     	Response.write "Transferencia entre bodegas"
			     Case 203
			     	Response.write txtODPIReq
			     Case 204
			     	Response.write txtODPIInv
			     Case 22
			     	Response.write txtOpor
			     End Select 
			     If rs("DocNum") <> "" then response.write " #" & rs("DocNum")
			     If rs("PayLogNum") <> "" Then Response.Write "/" & txtRct %>&nbsp;</td>
				<td style="height: 15px"><% Select Case rs("Status")
			    Case "H"
			    	Response.write "" & getventasXLngStr("LtxtConf") & ""
			    Case "R" 
			    	Response.Write "" & getventasXLngStr("LtxtPend") & ""
			    End Select %>&nbsp;</td>
				<td style="height: 15px"><p align="right"><nobr><%=rs("Currency")%>&nbsp;<%=FormatNumber(rs("DocTotal"),myApp.SumDec)%></nobr></td>
			</tr>
			<% If rs("Comments") <> "" Then %>
			<tr class="GeneralTbl">
				<td width="15">
				&nbsp;</td>
				<td style="width: 18px">
				&nbsp;</td>
				<td colspan="<% If Request("cmd") <> "activeClient" Then %>11<% Else %>7<% End If %>"><%=getventasXLngStr("DtxtObservations")%>: <%=rs("Comments")%></td>
			</tr>
			<% End If %>
			 <% rs.movenext
			 loop  %>
			<tr class="GeneralTblBold2">
				<td colspan="13">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr class="GeneralTblBold2">
						<td><input type="submit" name="btnDel" value="<%=getventasXLngStr("DtxtDelete")%>"><input type="hidden" name="go2" value="<% If Not activeClient Then %>D<% Else %>AC<% End If %>"></td>
						<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">(<%=rx.recordcount%>) <%=getventasXLngStr("LtxtDocuments")%></td>
					</tr>
				</table>
				</td>
			</tr>
			<% Else %>
			<tr class="GeneralTblBold2">
				<td colspan="13">
				<p align="center"><%=getventasXLngStr("DtxtNoData")%></td>
			</tr>
			<% End If %>
		</table>
		</td>
	</tr>
	<% If rx.PageCount > 1 Then %>
	<tr>
		<td align="center"><% doVentasXPages %></td>
	</tr>
	<% End If %>
<% for each Item in Request.Form
If Item <> "chkDel" and Item <> "btnDel" and Item <> "go2" and Item <> "orden1" and Item <> "orden2" Then %><input type="hidden" name="<%=Item%>" value="<%=Request.Form(Item)%>"><% End If
next 

for each Item in Request.QueryString %><input type="hidden" name="<%=Item%>" value="<%=Request.QueryString(Item)%>"><% next %>
<% If Request("orden1") = "" Then %>
<input type="hidden" name="orden1" value="0">
<input type="hidden" name="orden2" value="D">
<% End If %>

</form>
</table>
</div>
<% Sub doVentasXPages %>
<table cellpadding="0" cellspacing="2" border="0" width="100%">
	<tr>
	<% If iCurNext > 1 Then %>
	<td width="14" class="FirmTlt3"><a href="javascript:goPage(<%= ((iCurNext-1)*15) %>);">
	<img border="0" src="design/0/images/<%=Session("rtl")%>prevAll.gif" width="12" height="13" align="left"></a>
	</td>
	<% End If %>
	<% If nPage > 1 Then %><td width="15" class="FirmTlt3">
	<p align="center">
	<a href="javascript:goPage(<%= nPage - 1 %>);">
	<img border="0" src="design/0/images/<%=Session("rtl")%>prev_icon_trans.gif" width="15" height="15"></a></td>
	<% End If %>
	<td class="FirmTlt3">
	<p align="center" dir="ltr"><% if rx.PageCount > 1 then
	For I = fromI To toI
		If I = nPage Then %>
			<font size="3">
			<b><%= I %></b></font>
			<% Else %>
			<a class="LnkSearchPaginacion" href="javascript:goPage(<%= I %>);"><%= I %></a>
			<% End If
	Next 'I
	end if %></td>
	<% If nPage < rx.PageCount Then %>
	<td width="15" class="FirmTlt3">
	<p align="center">
	<a href="javascript:goPage(<%= nPage + 1 %>);">
	<img border="0" src="design/0/images/<%=Session("rtl")%>next_icon_trans.gif" width="15" height="15"></a></td>
	<% End If %>
	<% If iCurNext < iCurMax Then %>
	<td width="14" class="FirmTlt3">
	  <a href="javascript:goPage(<%= (iCurNext*15)+1 %>);">
	  <img border="0" src="design/<%=SelDes%>/images/<%=Session("rtl")%>nextAll.gif" width="12" height="13" align="right"></a>
	</td>
	<% End If %>
	</tr>
</table>
<% End Sub %>
<% 
Sub doVentasXSortImg(c)
	If orden1 = c Then
		If orden2 = "A" Then
			Response.Write "<img src=""images/arrow_up.gif"">"
		Else
			Response.Write "<img src=""images/arrow_down.gif"">"
		End If
	End If
End Sub 
Sub doVentasXSortBG(c)
	If orden1 = c Then Response.Write "class=""GeneralTblBold2HighLight"""
End Sub %>
<script language="javascript">
function doCheckDel(Img, LogNum)
{
	if (!document.getElementById('chkDel' + LogNum).checked)
	{
		document.getElementById('chkDel' + LogNum).checked = true;
		Img.src = 'images/checkbox_on.jpg';
	}
	else
	{
		document.getElementById('chkDel' + LogNum).checked = false;
		Img.src = 'images/checkbox_off.jpg';
	}
}
function goPage(p) { document.frmGoX.page.value = p; document.frmGoX.submit(); }
function doSort(c)
{
	document.frmGoX.orden1.value = c;
	if ('<%=orden1%>' == c)
	{
		if ('<%=orden2%>' == 'A')
			document.frmGoX.orden2.value = 'D';
		else
			document.frmGoX.orden2.value = 'A';
	}
	else
	{
		document.frmGoX.orden2.value = 'A';
	}
	document.frmGoX.page.value = 1;
	document.frmGoX.submit();
}
function delDoc(LogNum)
{
	if(!confirm('<%=getventasXLngStr("LtxtCondDelDoc")%>'.replace('{0}', LogNum))) return;
	doMyLink('ventas/docdel.asp', 'retval='+LogNum+varx, '');
}
function doGoDoc(obj, logNum, payLogNum, CardCode, Status)
{
	if (Status == 'H') if (!confReopen()) return;
	
	if (obj == 24) document.frmGoDoc.action = 'payments/go.asp';
	else document.frmGoDoc.action = 'ventas/go.asp';
	document.frmGoDoc.doc.value = logNum;
	document.frmGoDoc.payDoc.value = payLogNum;
	document.frmGoDoc.cl.value = CardCode;
	document.frmGoDoc.status.value = Status;
	document.frmGoDoc.submit();
}
</script>
<form name="frmGoDoc" method="post" action="">
<input type="hidden" name="doc" value="">
<input type="hidden" name="payDoc" value="">
<input type="hidden" name="cl" value="">
<input type="hidden" name="status" value="">
</form>
<form name="frmGoX" method="post" action="<% If SearchCmd = "activeClient" Then %>activeClient<% Else %>searchOpenedDocs<% End If %>.asp">
<input type="hidden" name="page" value="">
<input type="hidden" name="retval" value="">
<% 
varx = ""
for each Item in Request.Form 
	If Item <> "retval" Then
	varx = varx & "&" & Item & "=" & Request.Form(Item)
	If Item <> "page" Then %>
	<input type="hidden" name="<%=Item%>" value="<%=Request.Form(Item)%>">
<%	End If
	End If
next 

for each Item in Request.QueryString
	If Item <> "retval" Then
	varx = varx & "&" & Item & "=" & Request.QueryString(Item) 
	If Item <> "page" Then %>
	<input type="hidden" name="<%=Item%>" value="<%=Request.QueryString(Item)%>">
<%	End If
	End If
next %>
<% If Request("orden1") = "" Then %>
<input type="hidden" name="orden1" value="0">
<input type="hidden" name="orden2" value="D">
<% End If %>
</form>
<script>
var varx = '<%=Replace(varx, "'", "\'")%>'
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
	document.frmViewDetail.DocType.value = DocType;
	document.frmViewDetail.submit();
}
</script>
<% Else %>
<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTlt">
		<td><%=getventasXLngStr("LttlPendDocs")%></td>
	</tr>
	<tr class="GeneralTblBold2">
		<td>
		<p align="center"><%=getventasXLngStr("DtxtNoData")%></td>
	</tr>
</table>
<% End If %>