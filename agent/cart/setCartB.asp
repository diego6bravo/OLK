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
<!--#include file="../clearItem.asp"-->
<!--#include file="lang/setCartB.asp" -->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<!--#include file="../myHTMLEncode.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getsetCartBLngStr("LttlCartLineBatch")%></title>
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css">
<%
set rs = Server.CreateObject("ADODB.recordset")

If Request.Form("btnSubmit") <> "" Then
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetItemBatchSaveList" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LogNum") = Session("RetVal")
	cmd("@LineNum") = CInt(Request("LineNum"))
	cmd("@ItemCode") = Request("Item")
	cmd("@WhsCode") = Request("WhsCode")
	set rs = cmd.execute()

	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "OLKSaveItemBatchData"
	cmd.Parameters.Refresh()
	cmd("@LogNum") = Session("RetVal")
	cmd("@LineNum") = CInt(Request("LineNum"))
		
	do while not rs.eof
		BatchNum = Replace(rs("BatchNum"), " ", "<:-Space->")
		SelQty = Request("SelQty" & BatchNum)
		cmd("@BatchNum") = rs("BatchNum")
		cmd("@Qty") = SelQty
		cmd.execute()
		rs.movenext
	loop
	doClose
Else
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetLineBatchDetails" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@LogNum") = Session("RetVal")
cmd("@LineNum") = CInt(Request("LineNum"))
cmd("@ItemCode") = Request("Item")
set rs = cmd.execute()

lineQty = CDbl(rs("Quantity"))
lineUnit = rs("SaleType")
WhsCode = rs("WhsCode")
SalPackUn = CDbl(rs("SalPackUn"))
NumInSale = CDbl(rs("NumInSale"))

Select Case CInt(lineUnit)
	Case 1
		ReqQty = lineQty
	Case 2
		ReqQty = lineQty*NumInSale
	Case 3
		ReqQty = lineQty*NumInSale*SalPackUn
End Select

If Request("order") = "" Then order = "OIBT.BatchNum" Else order = saveHTMLDecode(Request("order"), True)
If Request("orderBy") = "" Then orderBy = "asc" Else orderBy = Request("orderBy")

%>
<script language="javascript" src="../general.js"></script>
</head>

<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" <% If Request("order") <> "" Then %>onload="SumTotal();"<% End If %>>
<script language="javascript">
function chkQty(fld, SelQty, MaxQty)
{
	if (!IsNumeric(fld.value) || fld.value == '')
	{
		alert('<%=getsetCartBLngStr("DtxtValNumVal")%>');
		fld.value = SelQty;
		fld.focus();
	}
	else if (parseFloat(fld.value) < 0)
	{
		alert('<%=getsetCartBLngStr("DtxtValNumMinVal")%>'.replace('{0}', '0'));
		fld.value = SelQty;
		fld.focus();
	}
	else if (parseFloat(fld.value) > parseFloat(MaxQty))
	{
		alert('<%=getsetCartBLngStr("LtxtValMoreThenAvl")%>');
		fld.value = MaxQty;
		fld.focus();
	}
	SumTotal();
}

function SumTotal()
{
	selQty = document.frmMain.SelQty;
	var sQtyCount = 0;
	if (selQty.length)
	{
		for (var i = 0;i<selQty.length;i++)
		{
			sQtyCount += parseFloat(selQty(i).value);
		}
	}
	else
	{
		sQtyCount = parseFloat(selQty.value);
	}
	document.frmMain.txtSelQty.value = sQtyCount;
	document.frmMain.txtOpenQty.value = parseFloat(document.frmMain.txtReqQty.value)-sQtyCount ;
}

function IsNumeric(sText)
{
   var ValidChars = "0123456789.-";
   var IsNumber=true;
   var Char;

 
   for (i = 0; i < sText.length && IsNumber == true; i++) 
      { 
      Char = sText.charAt(i); 
      if (ValidChars.indexOf(Char) == -1) 
         {
         IsNumber = false;
         }
      }
   return IsNumber;
   
}
function valFrm()
{
	if (parseFloat(document.frmMain.txtOpenQty.value) < 0)
	{
		alert('<%=getsetCartBLngStr("LtxtValReqQty")%>');
		return false;
	}
	return true;
}
function doSort(ColName)
{
	document.frmMain.order.value = ColName;
	document.frmMain.orderBy.value = 'asc';
	if ('<%=order%>' == ColName)
	{
		if ('<%=orderBy%>' == 'asc') document.frmMain.orderBy.value = 'desc';
	}
	document.frmMain.submit();
}
</script>
<table border="0" width="100%" id="table1">
	<form method="POST" action="setCartB.asp" name="frmMain" onsubmit="return valFrm();" webbot-action="--WEBBOT-SELF--">
	<tr>
		<td colspan="4" class="GeneralTblBold2"><%=getsetCartBLngStr("LttlBatchSel")%></td>
	</tr>
	<tr>
		<td class="GeneralTblBold2"><%=getsetCartBLngStr("DtxtItem")%>:</td>
		<td class="GeneralTbl"><%=Request("Item")%>&nbsp;</td>
		<td class="GeneralTblBold2"><%=getsetCartBLngStr("DtxtWarehouse")%>:</td>
		<td class="GeneralTbl"><%=rs("WhsName")%>&nbsp;</td>
	</tr>
	<tr>
		<td class="GeneralTblBold2"><%=getsetCartBLngStr("DtxtDescription")%>:</td>
		<td class="GeneralTbl" colspan="3"><%=rs("ItemName")%>&nbsp;</td>
	</tr>
	<tr>
		<td class="GeneralTblBold2"><%=getsetCartBLngStr("LtxtReqQty")%>:</td>
		<td class="GeneralTbl">
		<input class="GeneralTbl" type="text" name="txtReqQty" value="<%=ReqQty%>" size="1" style="text-align: right; width: 100%; border-style: solid; border-width: 0"></td>
		<td class="GeneralTblBold2">&nbsp;</td>
		<td class="GeneralTbl">&nbsp;</td>
	</tr>
	<tr>
		<td class="GeneralTblBold2"><%=getsetCartBLngStr("LtxtSelQty")%>:</td>
		<td class="GeneralTbl"><input class="GeneralTbl" type="text" name="txtSelQty" value="<%=rs("SelQty")%>" size="1" style="text-align: right; width: 100%; border-style: solid; border-width: 0"></td>
		<td class="GeneralTblBold2">&nbsp;</td>
		<td class="GeneralTbl">&nbsp;</td>
	</tr>
	<tr>
		<td class="GeneralTblBold2"><%=getsetCartBLngStr("LtxtPendQty")%>:</td>
		<td class="GeneralTbl"><input class="GeneralTbl" type="text" name="txtOpenQty" value="<%=ReqQty-CDbl(rs("SelQty"))%>" size="1" style="text-align: right; width: 100%; border-style: solid; border-width: 0"></td>
		<td class="GeneralTblBold2">&nbsp;</td>
		<td class="GeneralTbl">&nbsp;</td>
	</tr>
	<%
	set rx = Server.CreateObject("ADODB.RecordSet")
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetBatchRepRead" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	rx.open cmd, , 3, 1
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetLineBatch" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@LogNum") = Session("RetVal")
	cmd("@LineNum") = CInt(Request("LineNum"))
	cmd("@ItemCode") = Request("Item")
	cmd("@WhsCode") = WhsCode
	set rs = cmd.execute() %>
	<tr>
		<td colspan="4">
		<table border="0" cellspacing="2" width="100%" id="table2">
			<tr>
				<td class="GeneralTblBold2<%=doSortLinkHighLight("OIBT.BatchNum")%>" style="cursor: hand" onclick="javascript:doSort('OIBT.BatchNum');">
				<%=getsetCartBLngStr("LtxtBatch")%><%=doSortLink("OIBT.BatchNum")%></td>
				<% If Not rx.Eof Then
				do while not rx.eof %>
				<td class="GeneralTblBold2<%=doSortLinkHighLight(CStr(rx("rowName")))%>" <% If InStr(rx("rowName"), "'") = 0 Then %>style="cursor: hand" onclick="javascript:doSort('<%=Replace(myHTMLEncode(rx("rowName")), "'", "\'")%>');"<% End If %>><%=rx("rowName")%><%=doSortLink(CStr(rx("rowName")))%></td>
				<% rx.movenext
				loop
				rx.movefirst
				End If %>
				<td class="GeneralTblBold2<%=doSortLinkHighLight("AvlQty")%>" style="cursor: hand" onclick="javascript:doSort('AvlQty');">
				<%=getsetCartBLngStr("LtxtAvl")%><%=doSortLink("AvlQty")%></td>
				<td class="GeneralTblBold2"><%=getsetCartBLngStr("LtxtSelection")%></td>
			</tr>
			<% do while not rs.eof
			AvlQty = rs("AvlQty")
			BatchNum = Replace(rs("BatchNum"), " ", "<:-Space->") %>
			<tr>
				<td class="GeneralTbl"><%=rs("BatchNum")%>&nbsp;</td>
				<% If Not rx.Eof Then
				do while not rx.eof
				strVal = rs("ItemRep" & rx("rowIndex")) %>
				<td class="GeneralTbl"><% If Not IsNull(strVal) Then %><%=strVal%><% End If %></td>
				<% rx.movenext
				loop
				rx.movefirst
				End If %>
				<td class="GeneralTbl" align="right"><%=AvlQty%>&nbsp;</td>
				<td class="GeneralTbl" align="right">
				<input type="text" name="SelQty<%=BatchNum%>" id="SelQty" size="20" value="<% If Request("SelQty" & BatchNum) = "" Then %><%=rs("SelQty")%><% Else %><%=Request("SelQty" & BatchNum)%><% End If %>" style="width: 100%; text-align: right" onfocus="this.select();" onchange="chkQty(this, <%=rs("SelQty")%>, <%=AvlQty%>)"></td>
			</tr>
			<% rs.movenext
			loop %>
		</table>
		</td>
	</tr>
	<tr>
		<td colspan="4" class="GeneralTblBold2">
		<table border="0" cellspacing="0" width="100%" id="table3">
			<tr>
				<td><input type="submit" value="<%=getsetCartBLngStr("DtxtSave")%>" name="btnSubmit"></td>
				<td>
				<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
				<input type="button" value="<%=getsetCartBLngStr("DtxtCancel")%>" name="btnCancel" onclick="javascript:if(confirm('<%=getsetCartBLngStr("LtxtValCloseWin")%>'))window.close();"></td>
			</tr>
		</table>
		</td>
	</tr>
		<input type="hidden" name="LineNum" value="<%=Request("LineNum")%>">
		<input type="hidden" name="Quantity" value="<%=lineQty%>">
		<input type="hidden" name="SelUn" value="<%=lineUnit%>">
		<input type="hidden" name="WhsCode" value="<%=myHTMLEncode(WhsCode)%>">
		<input type="hidden" name="Item" value="<%=myHTMLEncode(Request("Item"))%>">
		<input type="hidden" name="order" value="<%=order%>">
		<input type="hidden" name="orderBy" value="<%=orderBy%>">
	</form>
</table>

</body>
<% 
rx.close
set rx = nothing
End If %>
<% 
conn.close
set rs = nothing %>
</html>
<% Sub doClose %>
<script>
<% 
If Request("txtSelQty") = 0 Then
	rCmd = 0
ElseIf Request("txtSelQty") = Request("txtReqQty") Then
	rCmd = 1
Else 
	rCmd = 2
End If
%>
opener.setSBImg(<%=rCmd%>);
window.close();
</script>
<% End Sub %>
<% Function doSortLink(ColName)
	retVal = ""
	If ColName = order Then
		If orderBy = "asc" Then
			sortByImg = "down"
		Else
			sortByImg = "up"
		End If
		retVal = "<img border=""0"" src=""../images/arrow_" & sortByImg & ".gif"">"
	End If
	doSortLink = retVal
End Function
Function doSortLinkHighLight(ColName)
	retVal = ""
	If ColName = order Then
		retVal = "HighLight"
	End If
	doSortLinkHighLight = retVal
End Function
%>