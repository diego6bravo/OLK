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
<!--#include file="lang/setCartS.asp" -->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<!--#include file="../myHTMLEncode.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getsetCartSLngStr("LttlCartLineSer")%></title>
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css">
<style>
.noborder {
	border-style:solid; border-width:0; background:background-image;
}
</style>
<%
set rs = Server.CreateObject("ADODB.recordset")

If Request.Form("btnSubmit") <> "" Then
	sql = "declare @LogNum int set @LogNum = " & Session("RetVal") & " " & _
		"declare @LineNum int set @LineNum = " & Request("LineNum") & " " & _
		"delete R3_ObsCommon..DOC4 where LogNum = @LogNum and LineNum = @LineNum "
		
	If Request("SerialNum") <> "" Then
		arrSer = Split(Request("SerialNum"), ", ")
		For i = 0 to UBound(arrSer)
			If Request("ChkSer" & arrSer(i)) = "Y" Then
				sql = sql & "insert R3_ObsCommon..DOC4(LogNum, LineNum, LineNum2, SysSerial) values(@LogNum, @LineNum, " & i & ", " & arrSer(i) & ") "
			End If
		Next
	End If
	conn.execute(sql)
	doClose
Else
sql = "select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', T0.ItemCode, T0.ItemName) ItemName, " & _
	"IsNull((select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OWHS', 'WhsName', T1.WhsCode, WhsName) from OWHS where WhsCode = T1.WhsCode collate database_default), '') WhsName, " & _
	"T0.SalPackUn, T0.NumInSale, T1.WhsCode, T2.SaleType, T1.Quantity/Case T2.SaleType When 3 Then T0.SalPackUn Else 1 End Quantity, " & _
	"IsNull((select Count('') from R3_ObsCommon..doc4 where LogNum = " & Session("RetVal") & " and LineNum = " & Request("LineNum") & "),0) SelQty " & _
	"from OITM T0 " & _
	"inner join R3_ObsCommon..DOC1 T1 on T1.LogNum = " & Session("RetVal") & " and T1.LineNum = " & Request("LineNum") & " " & _
	"inner join OLKSalesLines T2 on T2.LogNum = T1.LogNum and T2.LineNum = T1.LineNum " & _
	"where T0.ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "'"
set rs = conn.execute(sql)

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
SelQty = CDbl(rs("SelQty"))
If Request("order") <> "" and Request("txtSelQty") <> "" Then SelQty = CInt(Request("txtSelQty"))

If Request("order") = "" Then order = "OSRN.MnfSerial" Else order = saveHTMLDecode(Request("order"), True)
If Request("orderBy") = "" Then orderBy = "asc" Else orderBy = Request("orderBy")

%>
<script language="javascript" src="../general.js"></script>
</head>

<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" <% If Request("order") <> "" Then %>onload="SumTotal();"<% End If %>>
<script language="javascript">
var varReqQty = <%=ReqQty%>;
var varSelQty = <%=SelQty%>;

function checkSum(fld)
{
	var sumQty = fld.checked ? 1 : -1;

	varSelQty = varSelQty + sumQty;

	document.frmMain.txtSelQty.value = varSelQty;
	document.frmMain.txtOpenQty.value = varReqQty - varSelQty;

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
		alert('<%=getsetCartSLngStr("LtxtValReqQty")%>');
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
<table border="0" width="100%">
	<form method="POST" action="setCartS.asp" name="frmMain" onsubmit="return valFrm();" webbot-action="--WEBBOT-SELF--">
	<tr>
		<td colspan="4" class="GeneralTblBold2"><%=getsetCartSLngStr("LttlSerialSel")%></td>
	</tr>
	<tr>
		<td class="GeneralTblBold2"><%=getsetCartSLngStr("DtxtItem")%>:</td>
		<td class="GeneralTbl"><%=Request("Item")%>&nbsp;</td>
		<td class="GeneralTblBold2"><%=getsetCartSLngStr("DtxtWarehouse")%>:</td>
		<td class="GeneralTbl"><%=rs("WhsName")%>&nbsp;</td>
	</tr>
	<tr>
		<td class="GeneralTblBold2"><%=getsetCartSLngStr("DtxtDescription")%>:</td>
		<td class="GeneralTbl" colspan="3"><%=rs("ItemName")%>&nbsp;</td>
	</tr>
	<tr>
		<td class="GeneralTblBold2"><%=getsetCartSLngStr("LtxtReqQty")%>:</td>
		<td class="GeneralTbl">
		<input class="GeneralTbl" type="text" name="txtReqQty" value="<%=ReqQty%>" size="1" style="text-align: right; width: 100%; border-style: solid; border-width: 0"></td>
		<td class="GeneralTblBold2">&nbsp;</td>
		<td class="GeneralTbl">&nbsp;</td>
	</tr>
	<tr>
		<td class="GeneralTblBold2"><%=getsetCartSLngStr("LtxtSelQty")%>:</td>
		<td class="GeneralTbl"><input class="GeneralTbl" type="text" name="txtSelQty" value="<%=SelQty%>" size="1" style="text-align: right; width: 100%; border-style: solid; border-width: 0"></td>
		<td class="GeneralTblBold2">&nbsp;</td>
		<td class="GeneralTbl">&nbsp;</td>
	</tr>
	<tr>
		<td class="GeneralTblBold2"><%=getsetCartSLngStr("LtxtPendQty")%>:</td>
		<td class="GeneralTbl"><input class="GeneralTbl" type="text" name="txtOpenQty" value="<%=ReqQty-SelQty%>" size="1" style="text-align: right; width: 100%; border-style: solid; border-width: 0"></td>
		<td class="GeneralTblBold2">&nbsp;</td>
		<td class="GeneralTbl">&nbsp;</td>
	</tr>
	<%
	set rx = Server.CreateObject("ADODB.RecordSet")
	sql = "select IsNull(alterRowName, rowName) rowName, rowType, rowTypeDec, rowTypeRnd, rowField " & _
	"from OLKSerialRep T0 " & _
	"left outer join OLKSerialRepAlterNames T1 on T1.rowIndex = T0.rowIndex and T1.LanID = " & Session("LanID") & " " & _
	"where T0.rowActive = 'Y' " & _
	"order by T0.rowOrder asc"
	rx.open sql, conn, 3, 1
	
	If rx.recordcount > 0 Then
		do while not rx.eof
		
		  If rx("rowType") = "L" or rx("rowType") = "M" or rx("rowType") = "H" Then
 			Select Case rx("rowTypeDec")
				Case "S"
					myDec = myApp.SumDec
				Case "P"
					myDec = myApp.PriceDec
				Case "R"
					myDec = myApp.RateDec
				Case "Q"
					myDec = myApp.QtyDec
				Case "%"
					myDec = myApp.PercentDec
				Case "M"
					myDec = myApp.MeasureDec
 			End Select
		  End If
		  
		  If rx("rowType") = "T" Then
			AddCode1 = ""
			AddCode2 = ""
		  Else
			AddCode1 = "OLKCommon.dbo.DBOLKCode" & Session("ID") & "('" & rx("rowType") & "', "
			AddCode2 = ", " & myDec & ")"
		  End If
		  
		  If rx("rowTypeRnd") = "Y" Then 
			rowTypeRnd1 = "Convert(Char(1),Convert(int,(10 * rand())))+ + Convert(nvarchar(20),("
			rowTypeRnd2 = "))"
		  Else
			rowTypeRnd1 = ""
			rowTypeRnd2 = ""
		  End If

		  rowQuery = rx("rowField")
		  
		  AddFields = AddFields & AddCode1 & "(" & rowTypeRnd1 & rowQuery & rowTypeRnd2 & ")" & AddCode2 & " As '" & Replace(rx("rowName"), "'", "''") & "', "
		rx.movenext
		loop
		rx.movefirst
	End If
	
	AddFields = Replace(AddFields, "@ItemCode", "OSRQ.ItemCode")
	AddFields = Replace(AddFields, "@SysNumber", "OSRN.BatchNum")
	AddFields = Replace(AddFields, "@WhsCode", "OSRQ.WhsCode")
	
	sql = "declare @LanID int set @LanID = " & Session("LanID") & " " & _  
		"select T0.SysNumber, OSRN.MnfSerial Serial, " & AddFields & _
		"CASE When T1.LogNum is not null Then 'Y' Else 'N' End Checked " & _  
		"FROM  OSRQ T0   " & _  
		"inner join OSRN on OSRN.ItemCode = T0.ItemCode and OSRN.SysNumber = T0.SysNumber " & _  
		"left outer join R3_ObsCommon..DOC4 T1 on T1.LogNum = " & Session("RetVal") & " and T1.LineNum = " & Request("LineNum") & " and T1.SysSerial = T0.SysNumber " & _  
		"WHERE T0.ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "' and T0.WhsCode = N'" & saveHTMLDecode(WhsCode, False) & "' and T0.Quantity <> 0 " & _  
		"order by " & order & " " & orderBy
	set rs = conn.execute(sql) %>
	<tr>
		<td colspan="4">
		<table border="0" cellspacing="2" width="100%" id="table2">
			<tr>
				<td class="GeneralTblBold2<%=doSortLinkHighLight("OSRN.MnfSerial")%>" style="cursor: hand" onclick="javascript:doSort('OSRN.MnfSerial');">
				<%=getsetCartSLngStr("LtxtSerialNum")%><%=doSortLink("OSRN.MnfSerial")%></td>
				<% If Not rx.Eof Then
				do while not rx.eof %>
				<td class="GeneralTblBold2<%=doSortLinkHighLight(CStr(rx("rowName")))%>" <% If InStr(rx("rowName"), "'") = 0 Then %>style="cursor: hand" onclick="javascript:doSort('<%=Replace(myHTMLEncode(rx("rowName")), "'", "\'")%>');"<% End If %>><%=rx("rowName")%><%=doSortLink(CStr(rx("rowName")))%></td>
				<% rx.movenext
				loop
				rx.movefirst
				End If %>
			</tr>
			<% do while not rs.eof
			SerialNum = rs("SysNumber") %>
			<tr>
				<td class="GeneralTbl"><input type="hidden" name="SerialNum" value="<%=SerialNum%>"><input type="checkbox" name="ChkSer<%=SerialNum%>" id="ChkSer<%=SerialNum%>" size="20" class="noborder" value="Y" <% If rs("Checked") = "Y" or Request("order") <> "" and Request("ChkSer" & SerialNum) = "Y" Then %>checked<% End If %> onfocus="this.select();" onclick="checkSum(this)"><label for="ChkSer<%=SerialNum%>"><%=rs("Serial")%></label></td>
				<% If Not rx.Eof Then
				do while not rx.eof %>
				<td class="GeneralTbl"><% If Not IsNull(rs(CStr(rx("rowName")))) Then %><%=rs(CStr(rx("rowName")))%><% End If %></td>
				<% rx.movenext
				loop
				rx.movefirst
				End If %>
			</tr>
			<% rs.movenext
			loop %>
		</table>
		</td>
	</tr>
	<tr>
		<td colspan="4" class="GeneralTblBold2">
		<table border="0" cellspacing="0" width="100%">
			<tr>
				<td><input type="submit" value="<%=getsetCartSLngStr("DtxtSave")%>" name="btnSubmit"></td>
				<td>
				<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
				<input type="button" value="<%=getsetCartSLngStr("DtxtCancel")%>" name="btnCancel" onclick="javascript:if(confirm('<%=getsetCartSLngStr("LtxtValCloseWin")%>'))window.close();"></td>
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
opener.setSSImg(<%=rCmd%>);
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